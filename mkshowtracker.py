#! /usr/bin/python3

#Programmed with python 3.8.9
#Ananlyzes a progress report and makes a tracker for Ronnie Perez

import os
import re
import sys
import openpyxl


def getXLF():

    print("")
    while True:
        xlf = input("  Give me your progress report (.xlsx extension): ")

        m = re.search("\.xlsx$",xlf)
        if m and os.path.exists(xlf):
            break
    return xlf

def selectWS(wb,xlf):

    print("")
    while True:

        print("  Work Sheets in", xlf)
        print("  ==============")
        for sheet in wb.sheetnames:
            print(" ",sheet)
        sheet = input("\n  Select Your Work Sheet : ")
        if sheet in wb.sheetnames:
            break
    print("")
    return sheet

####Main Program####
xlf = getXLF()

try:
    wb = openpyxl.load_workbook(filename=xlf, read_only=True)
except:
    print("  ~~Could not open", xlf,"\n")
    sys.exit(1)

sheet = selectWS(wb,xlf)
ws = wb[sheet]

MAX_NONE = 50   #maximum number of None (assumes we get to the end)
RECORD   = False

r         = 1
showcol   = 1
epcol     = 2
nonecount = 0
showdict  = {}

#Get the data
print("Analyzing",end="")
while True:
    show = str(ws.cell(row=r,column=showcol).value)
    snum = str(ws.cell(row=r,column=epcol).value)

    if show == 'None':
        nonecount += 1
        if nonecount == MAX_NONE:
            break
    else:
        season = show + ':' + snum.zfill(2)
        if season in showdict:
            n = showdict[season]
            n += 1
            showdict[season] = n
        else:
            if show != 'Show Title':
                showdict[season] = 1
    r += 1
    print(".",end="")

print("")
for s in sorted(showdict.keys()):
    print(s)

wb.close()
