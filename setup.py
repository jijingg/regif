#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
__author__ = "jijing"

from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(name='regif',
      version='1.3.3', 
      description='register interface verilog code generator with leggal check',  
      author='jijing.guo',                     
      author_email='goco.v@163.com',           
      url='https://github.com/jijingg/regif',  
      packages=find_packages(),           
      long_description=long_description,  
      long_description_content_type="text/markdown",   
      license="GPLv3",   
      classifiers=[
          "Programming Language :: Python :: 3", 
          "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
          "Operating System :: OS Independent"],

      python_requires='>=3.3',   
      install_requires=[
          "xlrd>=1.1.0",
          "python_docx>=1.11.3",
          ]
      )

