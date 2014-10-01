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
    reg = r',([^,]*' + tname + r'[^,]*),'
    try:
        page1 = int(npl[k]['pages'][0])
        tpage1 = str(pdflist[page1-1]).replace('\n',',')
        exps1 = re.findall(reg,tpage1)
        try:
            page2 = int(npl[k]['pages'][1])
            tpage2 = str(pdflist[page2-1]).replace('\n',',')
            exps2 = re.findall(reg,tpage2)
            try:
                pi.cell(row = k+2, column = 9).value = exps1[0]
            except:
                ok
            try:
                pi.cell(row = k+2, column = 10).value = exps2[0]
            except:
                ok
            try:
                pi.cell(row = k+2, column = 11).value = str(exps1[1:])
            except:
                ok
            try:
                pi.cell(row = k+2, column = 12).value = str(exps2[1:])
            except:
                ok
        except:
            try:
                pi.cell(row = k+2, column = 9).value = exps1[0]
            except:
                ok
            try:
                pi.cell(row = k+2, column = 10).value = exps1[1]
            except:
                ok
            try:
                pi.cell(row = k+2, column = 11).value = exps1[2]
            except:
                ok
            try:
                pi.cell(row = k+2, column = 12).value = str(exps1[3:])
            except:
                ok
        print tname + ': ' + exps1[0]
    except:
        print tname
        

pi9.save('PI_2009_Updated.xlsx')
    
#t1 = pdflist[257]
#t1 = str(t1)
#    .{0,15}Aaberg.{0,15}
#re.findall(r',([^,]*Aaberg[^,]*),',tpage)