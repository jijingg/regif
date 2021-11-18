#!/usr/bin/env python
# coding=utf-8
# Date   : 2017.7.28
# Author : jijing.guo (All rights reserved) 
# Email  : goco.v@163.com
# version: 0.10(realize RW/RO/WO/RWW/RWC/RWT feature)
#        : 1.00(code refactor, add RC/RWP/P(Protect) feature)
#        : 1.10(add word.docx generate feature 2018.3.27)
#        : 1.20(add keword FIFO_IF)
#        : 1.30(add ClockGate on RW/RWT/WO reg)
#        : 1.31(argparser supported)
#        : 1.32(add read Group feature)
#        : 1.33(add reset-macro)

import sys, os, json, shutil 
import logging, re

def checklib():
    info = ""
    try:
        import xlrd
    except ImportError as e:
        info += str(e) + ", 'pip install xlrd' first!! \n"
    # try:
    #     from docx import Document
    # except ImportError as e:
    #     info += str(e) + ", 'pip install python_docx' first!! \n"
    if(info):
        info = "Error: \n" + info 
        print(info)
        exit(0)
    
# dependency load
checklib()
import xlrd
# from docx import Document

# logging 
class CustomFormatter(logging.Formatter):
    """Logging Formatter to add colors and count warning / errors"""
    grey     = "\x1b[38;1m"
    yellow   = "\x1b[33;1m"
    red      = "\x1b[31;1m"
    bold_red = "\x1b[31;1m"
    reset    = "\x1b[0m"
    format = "\n%(levelname)s - %(message)s"
    # format = "\n%(levelname)s - %(message)s (%(filename)s:%(lineno)d)"

    FORMATS = {
        logging.DEBUG:    grey     + format + reset,
        logging.INFO:     grey     + format + reset,
        logging.WARNING:  yellow   + format + reset,
        logging.ERROR:    red      + format + reset,
        logging.CRITICAL: bold_red + format + reset
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)

log = logging.getLogger("regif")
log.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setFormatter(CustomFormatter())
log.addHandler(ch)

# common
reg_format = [
    ("offset ", r"0x[\da-fA-F]+\Z", '"0x001C"'),
    ("regname", r"[a-zA-Z]\w*\Z", '"M_DMA_CFG"'),
    ("reg_des", r".*", 'some description is better '),
    ("width  ", r"[1-9]\d{1,2}\Z", '1~999 is better')]

field_format = [
    ("section", r"\[(\d+|\d+\:\d+)\]\Z", '"[0]" or "[15:7]" '),
    ("fd_name", r"[a-zA-Z]\w*\Z", 'must be one word '),
    ("RW     ", r"(RW|RC|BP|CP|RO|WO|RWW|RWC|RWT|RWP)\Z", "RW,RC,BP,CP,RO,WO,RWW,RWC,RWT,RWP only"),
    ("reset  ", r"(\d+'(h|H)[\da-fA-F]+\Z|\d+'(d|D)\d+\Z|\d+'(b|B)[01]+\Z|0\Z)", '''"0","3'd7", "12'hFFF", "2'b11"'''),
    ("wprotect", r"BITWISE|bitwise|\[(\d+|\d+\:\d+)\]\Z|\A\Z", "'[0]' '[12:8]' 'bitwise' 'BITWISE' or empty is ok"),
    ("locksig", r"[a-zA-Z]\w*\Z|\A\Z", "empty or one word be allowd"),
    ("fd_des ", r".*", ""),
    ("other  ", r".*", "")]

other_format = [
    (1, 1, "moudle ", r"[a-zA-Z]\w*", "first word need be a valid word"),
    (3, 5, "addr_wdith ", r"([2-9]\Z|[1-9]\d{1,2}\Z)", "2~999 is better"),
    (4, 5, "dw ", r"([4-9]\Z|[1-9]\d{1,2}\Z)", "4~999 is better"),
    (6, 5, "reg_pre ", r'\"[a-zA-Z_]\w*\"\Z|\"\"\Z', '"if_","reg_","" are better')]

# template 
ALWAYS_BEGIN = '''always @(posedge bmi_clk or negedge bmi_rstn)
    if(!bmi_rstn) begin '''
ALWAYS_BEGIN_MACRO="`ctrl_always_ff_begin(bmi_clk, bmi_rstn)"

RESET_MACRO="""
`ifndef __reset_macro_vh__
  `define __reset_macro_vh__
  `ifdef CTRL_RESET_ASYNC_LOW
    `define ctrl_always_ff_begin(clk,rst) always @(posedge clk or negedge rst) if(!rst) begin
  `elsif CTRL_RESET_ASYNC_HIGH
    `define ctrl_always_ff_begin(clk,rst) always @(posedge clk or posedge rst) if(rst) begin 
  `elsif CTRL_RESET_SYNC_LOW
    `define ctrl_always_ff_begin(clk,rst) always @(posedge clk) if(!rst) begin
  `elsif CTRL_RESET_SYNC_HIGH
    `define ctrl_always_ff_begin(clk,rst) always @(posedge clk) if(rst) begin
  `else             
    `define ctrl_always_ff_begin(clk,rst) always @(posedge clk or negedge rst) if(!rst) begin
  `endif   
`endif
"""

# common function
get_regif_dir = lambda : os.path.split(os.path.realpath(__file__))[0]

def trystr(value, ignorecase=0):
    _tmp = str(int(value)) if type(value) == float else \
        str(value) if type(value) == int else value
    return _tmp.lower() if ignorecase else _tmp

def tryint(value):
    s = trystr(value)
    return int(s)

def trybool(value, expect):
    return True if (trystr(value).upper() == expect) else False

def format_check(strs, line, row, info):
    if not re.match(info[1], strs):
        wstrs = '"' + strs + '"'
        return "Error: position {:3}{}: {:28} format error! Suggest like : {}\n" \
            .format(line + 10, chr(65 + row), wstrs, info[2])
    else:
        return ""

def title_check(bs):
    msg = ""
    for formats in other_format:
        line = formats[0]
        row = formats[1]
        info = formats[2:]
        msg += format_check(trystr(bs.cell_value(line, row)), line, row, info)
    return msg

def split2fixwidth(strlists, width):
    _out = "\n    "
    _len = 0
    for _str in strlists:
        _step = (len(_str) + 2)
        _len += _step
        if _len > width and (width - (_len - _step)) < 8:
            _len = 0
            _out += ("\n    " + _str + ", ")
        else:
            _out += (_str + ", ")
    return _out

def ispow2(x):
    return False if x & (x-1) else True

def groupSize(info):
    if info:
        gs = int(info)
        if(ispow2(gs)):
            return gs
        else:
            log.error("%d is not pow2, 32, 64, 128 is recommended, Fix at '5K' postion, then re-run again\n"%gs)
            exit()
    else:
        return 0

def getopt(args):
    withopt = [x.strip("opt=") for x in args if "opt=" in x]
    if withopt:
        return withopt[0]
    else :
        return ""

def intger(s: str):
    if(s.startswith("0x")):
        return int(s, 16)
    elif(s.startswith("0o")):
        return int(s, 8)
    elif(s.startswith("0b")):
        return int(s, 2)
    else :
        return int(s, 10)

# excel load and parser
class XLSParser():
    _regs = []
    _js = {}
    def __init__(self, bs):
        self._static_check(bs)
        self.getjson()
        self.modulename = re.match(other_format[0][3], bs.cell(2, 1).value).group()
        self.aw = tryint(bs.cell(3, 5).value)
        self.dw = tryint(bs.cell(4, 5).value)
        self.amsb = int(self.aw) - 1
        self.dmsb = int(self.dw) - 1
        self.withcg     = trybool(bs.cell(3, 10).value, "YES")
        self.reg_pre    = trystr(bs.cell(6, 5).value).strip('"')
        self.groupsize  = groupSize(bs.cell(4, 10).value)
        self.unlocktype = ("RO", "RC", "RWT", "RWP", "CP", "BP")
        self.lock_pre   = "o"
        self.page_width = 50
        self._hw_clc    = "_hw_clc"
        self._hw_set    = "_hw_set"
        self._hw_setval = "_hw_setval"
        self.always_ff_begin = ALWAYS_BEGIN_MACRO if(bs.cell(5,10).value.upper() == "YES") else ALWAYS_BEGIN
        self.macros = RESET_MACRO if(bs.cell(5,10).value.upper()=="YES") else ""
        self.regs = json2reg(self._js)
        # self._dynamic_check()

    def __repr__(self) -> str:
        return "\n".join([str(s) for s in self.regs])

    def getjson(self):
        def sec(s):
            ns = str2sec(s)
            return {"msb" : ns.msb, "lsb" : ns.lsb}

        def field(t):
            return {"sec": sec(t[0]), "name":t[1], "acc" : t[2],  "reset" : trystr(t[3]), "wp": t[4], "lock": t[5], "doc":t[6]}

        def reg(x):
            regname = x[1]
            fields = [field(t) for t in x[4]]
            return { regname : {
                "offset" : intger(x[0]),
                "doc" : x[2],
                "fields" : fields}
            }

        def tryupdate(x):
            key = list(x.keys())[0]
            if key in self._js:
                log.error("{} already exits".foramt(key))
                exit(2)
            self._js.update(x)

        [tryupdate(reg(x)) for x in self._regs]

    def getregs(self, bs):
        nrows = bs.nrows
        ncols = bs.ncols
        address_idx = 0
        for i in range(9, nrows):
            row_data = bs.row_values(i)
            if row_data[0] != u'':
                self._regs.append(row_data[0:4])
                self._regs[address_idx].append([])
                self._regs[address_idx][4].append(row_data[4:ncols])
                address_idx += 1
            else:
                self._regs[address_idx - 1][4].append(row_data[4:ncols])

    def _dynamic_check(self):
        print("Dynamic check ..... ")
        msg = ""
        for reg in self.regs:
            msg += reg.dynamic_check()
        if msg:
            msg += "\nFixed all those format Error first, then regenerate again !\n"
            print(msg)
            exit(0)
        else:
            print("Dynamic check Pass! \n")
        return

    def _static_check(self, bs):
        self.getregs(bs)
        print("Static check ..... ")
        msg = title_check(bs)
        msg += static_check(self._regs)
        if msg:
            msg += "\nFixed all those format Error first, then regenerate again !\n"
            print(msg)
            exit(0)
        else:
            print("Static check Pass! \n")
        return

    def show(self):
        print(self._js)

    def dumpjson(self):
        content = json.dumps(self._js, sort_keys=False, indent=4)
        with open("{}.json".format(self.modulename), "w") as f:
            f.write(content)

def json2reg(js):
    return [Reg(name, js[name]) for name in js]


def static_check(regs):
    msg = ""
    line = 0
    regs_unq_msg, addrs_unq_msg, fields_unq_msg = ("",) * 3
    regs_unq, addrs_unq, fields_unq = ([],) * 3
    for reg in regs:
        if reg[1] in addrs_unq:
            addrs_unq_msg += "Error: position {:3}{}: Regname {:20} already exist\n".format(line + 10, chr(65 + 1), reg[1])
        if reg[0] in regs_unq:
            regs_unq_msg += "Error: position {:3}{}: Addrs   {:20} already exist\n".format(line + 10, chr(65 + 0), reg[0])
        addrs_unq.append(reg[1])
        regs_unq.append(reg[0])
        for i in range(len(reg_format)):
            strs = trystr(reg[i])
            msg += format_check(strs, line, i, reg_format[i])
        for field in reg[4]:
            if field[1] in fields_unq and field[1] != "RESERVED":
                fields_unq_msg += "Error: position {:3}{}: Field   {:20} already exist\n".format(line + 10, chr(65 + 4), field[1])
            fields_unq.append(field[1])
            for i in range(len(field_format)):
                strs = trystr(field[i])
                msg += format_check(strs, line, i + 4, field_format[i])
            line += 1
    return msg + regs_unq_msg + addrs_unq_msg + fields_unq_msg

# object 
class Wire:
    dc_name = "wire"
    def __init__(self, name, dw = 1):
        self.name = name 
        self.dw = dw 
        self.align = 20

    def __str__(self):
        ndw = "" if(self.dw == 1) else "[{}:0]".format(self.dw-1)
        ndw = ndw.center(10)
        nname = self.name.ljust(self.align)
        dcname = self.dc_name.ljust(10)
        return  "{} {} {}".format(dcname, ndw, nname)
                                                 
    def __repr__(self):
        return self.__str__()

class Input(Wire):
    dc_name = "input"

class Output(Wire):
    dc_name = "output"
 
class OutputReg(Wire):
    dc_name = "output reg"

bmiport = [
    Input("bmi_clk"),
    Input("bmi_rstn"),
    Input("bmi_rd"),
    Input("bmi_wr"),
    Input("bmi_addr", 16),
    Input("bmi_wdata", 32),
    Output("bmi_rdata", 32),
    Output("bmi_rdvld"),
]

class Declars:
    jdot = ";\n"
    def __init__(self, signals):
        self.maxwidth = max([len(s.name) for s in signals])
        self.signals = map(self.align, signals)

    def align(self, s):
        s.align = self.maxwidth + 2
        return s

    def _joind(self):
        return self.jdot.join(map(str, self.signals))

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        return self._joind() + ";"

class IODeclars(Declars):
    jdot = ",\n"
    def __str__(self):
        return self._joind()
        # return super().__str__() + ";"     

class InstPort():
    def __init__(self, pname, wname, io, dw = 1):
        self.port = pname
        self.wire = wname
        self.io = io
        self.pgap = 30
        self.wgap = 30
        self.dw = dw
        self.islast = False
                                          
    def __str__(self):
        pname = self.port.ljust(self.pgap)
        dot = " " if self.islast else ","
        io = self.io
        ndw = "" if(self.dw == 1) else "[{}:0]".format(self.dw-1)
        wname = (self.wire+ndw).ljust(self.wgap+10)
        return  "    .{} ( {} ){}//{}".format(pname, wname, dot, io)

class InstPorts():
    def __init__(self, ports):
        self.pmax = max([len(p.port) for p in ports])
        self.wmax = max([len(p.wire) for p in ports])
        self.ports = list(map(self.align, ports))
        self.ports[-1].islast = True

    def align(self, s):
        s.pgap = self.pmax + 2
        s.wgap = self.wmax + 2
        return s

    def __str__(self):
        return "\n".join(map(str, self.ports))
    def __repr__(self):
        return self.__str__()  


class Reg:
    "name/offset/doc"
    def __init__(self, name, value):
        self.name = name
        self.offset = value["offset"]
        self.doc = value["doc"]
        self.fields = [Field(fd) for fd in value["fields"]]

    def __repr__(self) -> str:
        return self.__str__()
    def __str__(self) -> str:
        return "reg:{} {} [{}]".format(self.name, self.offset, "|".join([str(s) for s in self.fields]))

class Field:
    "sec/name/acc/reset/wp/lock/doc"
    def __init__(self, fieldjs):
        'deep copy is very important unless it will silence change fieldjs dict'
        import copy
        js = copy.deepcopy(fieldjs)
        js["sec"] = js2sec(js["sec"])
        js["reset"] = resetvalue(js["reset"])
        self.__dict__.update(js)

    def __repr__(self) -> str:
        return self.__str__()
    def __str__(self) -> str:
        return "filed:{}{}".format(self.name, self.sec)

def js2sec(js):
    return Section(js["msb"], js["lsb"])

def str2sec(s):
    multi_bit = re.match(r"\[(?P<b>\d+):(?P<s>\d+)\]", s)
    single_bit = re.match(r"\[(?P<b>\d+)\]", s)
    if multi_bit:
        msb = int(multi_bit.group(1))
        lsb = int(multi_bit.group(2))
        return Section(msb, lsb)
    elif single_bit:
        msb = int(single_bit.group(1))
        return Section(msb, msb)
    else:
        raise "bit section error"

def resetvalue(s):
    decm = re.match("(\d+)'d(\d+)", s)
    binm = re.match("(\d+)'b([01]+)", s)
    hexm = re.match("(\d+)'h([\da-f]+)", s)
    if s.isdigit():
        rel_val = int(s)
    elif decm:
        rel_val = int(decm.group(2))
    elif binm:
        rel_val = int(binm.group(2), 2)
    elif hexm:
        rel_val = int(hexm.group(2), 16)
    else:
        raise "An except error, the reset value not match 'd 'b 'h 0 type"
    return rel_val

class Section:
    def __init__(self, msb, lsb):
        self.msb = msb
        self.lsb = lsb
        self.width = msb - lsb + 1

    def __str__(self):
        return "" if(self.msb == self.lsb)  else "[{}:{}]".format(self.msb, self.lsb)
    def __repr__(self):
        return self.__str__()  

from enum import Enum
class Acc(Enum):
    RW = 1
    WO = 2
    RO = 3
    RC = 4
    WC = 5
    WT = 6
    WP = 7

# argvs process
def checkargs(args, info):
    if(len(args) == 0):
        log.error("misuage, '{}' try again".format(info))
        exit(2)

def creatxls(name):
    template = os.path.join(get_regif_dir(), "regif_template.xls")
    dest = "%s_regif.xls" % name
    if os.path.exists(dest):
        log.error("%s already exists !!!" % dest)
        exit(1)
    shutil.copy(template, dest)
    log.info("RegIf excel template %s created\n" % dest)

def task_regif(args):
    checkargs(args, "regif --init youmodule.xls")
    withdoc = "--doc" in args
    withjson = "--json" in args
    opt = getopt(args)
    xls = args[0]
    if (not os.path.exists(xls)):
        log.error("%s not found, check please, 'regif -h' for help" % xls)
        exit(2)
    wb = xlrd.open_workbook(xls)
    bs = wb.sheet_by_index(0)
    xlsreg = XLSParser(bs)
    xlsreg.dumpjson()
    print(xlsreg)

def task_creat(args):
    checkargs(args, "regif --init youmodule.xls")
    creatxls(args[1])

def task_ui(args):
    help = """
    regif --init mymodule            //creat template
    regif mymodule.xls               //generate .v
    regif mymodule.xls --doc         //generate .v .doc
    regif mymodule.xls --doc --json  //generate .v .doc .json
    """
    print(help)

task_entry = {
    "--init" : task_creat, 
    "xls"    : task_regif,
    "--help" : task_ui, 
    "-h"     : task_ui, 
}

def main():
    argvs = sys.argv[1:]
    noargs = not argvs
    if(noargs):
        task_ui([])
    else:
        command = argvs[0]
        args    = argvs
        command = "xls" if ".xls" in command else command
        task = task_entry.get(command, task_ui)
        task(args)

if __name__ == "__main__":
    try: 
        main()
    except:
        raise
