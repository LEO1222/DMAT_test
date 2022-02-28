#!/usr/bin/env python3
import mhi.pscad
import logging
import mhi.pscad.handler
import os, openpyxl
import pandas as pd

############# Print current directory #############
os.chdir("Automation scripts")
print ("CURRENT DIRECTORY:  %s" % os.getcwd())
###################################################

os.system("python ./large1_120.py")                     #1
os.system("python ./muyuan.py")                         #2
os.system("python ./yuxiang.py")                        #3
os.system("python ./ziheng_170_173blue.py")             #4
os.system("python ./ziheng_170_173grey.py")             #5
os.system("python ./ziheng_170_173orange.py")           #6
os.system("python ./ziheng_170_173yellow.py")           #7
os.system("python ./ziheng_174_177blue.py")             #8
os.system("python ./ziheng_174_177yellow.py")           #9
os.system("python ./ziheng_186_189blue.py")             #10
os.system("python ./ziheng_186_189green.py")            #11
os.system("python ./ziheng_186_189red.py")              #12
os.system("python ./ziheng_190_192.py")                 #13
os.system("python ./ziheng_193_198_40.py")              #14
os.system("python ./ziheng_193_198_60.py")              #15
os.system("python ./ziheng_193_198_minus_40.py")        #16
os.system("python ./ziheng_193_198_minus_60.py")        #17
os.system("python ./ziheng_step_response_155_160.py")   #18
os.system("python ./ziheng_step_response_178_181.py")   #19
os.system("python ./ziheng_step_response_182_185.py")   #20
os.system("python ./yuxiang_206FRT.py")                 #21