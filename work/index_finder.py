# -*- coding: utf-8 -*-
"""
Created on Wed Sep 24 14:29:17 2014

@author: sameer
"""

import re
from openpyxl import load_workbook

pitest = open('pilist.txt','r')
pdflist =eval(pitest.readlines()[0])
pitest.close()
  
l= 1
pi9 = load_workbook("/home/sameer/PI_2009.xlsx")
pi = pi9.active
npl = []
for rowno in range(2,4324):
    tempd = dict()
    tempd["lname"] = str(pi.cell(row = rowno, column = 2).value)
    tempd["pages"] = pi.cell(row = rowno, column = 8).value.split(',')
    npl.append(tempd)

for k in range(len(npl)):
    tname = npl[k]['lname']
    try:
        page = int(npl[k]['pages'][0])
        tpage = str(pdflist[page-1])
        reg = r'.{0,20}' + tname + r'.{0,20}'
        exps = re.findall(reg,tpage)
        pi.cell(row = k+2, column = 9).value = exps[0]
        try:
            pi.cell(row = k+2, column = 10).value = exps[1]
            pi.cell(row = k+2, column = 11).value = exps[2]
            pi.cell(row = k+2, column = 12).value = str(exps[3:])
        except:
            l +=1
        print tname + ': ' + exps[0]
    except:
        print tname
        

pi9.save('PI_2009_Updated.xlsx')
    
#t1 = pdflist[257]
#t1 = str(t1)
#    .{0,15}Aaberg.{0,15}
#re.findall(r',([^,]*Aaberg[^,]*),',tpage)