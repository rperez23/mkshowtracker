#! /usr/bin/python3

#Programmed with python 3.8.9
#Ananlyzes a progress report and makes a tracker for Ronnie Perez

import os
import re
import sys
import openpyxl

channeldict = {}

#get the name of the xlf (Progress Report)
def getXLF():

    print("")
    while True:
        xlf = input("  Give me your progress report (.xlsx extension): ")

        m = re.search("^(.+)\.xlsx$",xlf)
        if m and os.path.exists(xlf):
            prefix = m.group(1)
            break
    return xlf,prefix

#select the worksheet from the Progress Report
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

def readXLF(ws):

    MAX_NONE = 50   #maximum number of None (assumes we get to the end)
    #RECORD   = False

    r         = 1
    showcol   = 1
    epcol     = 2
    nonecount = 0
    showdict  = {}

    #Analyze the Progress Report
    print("Analyzing",end="",flush=True)
    while True:
        show = str(ws.cell(row=r,column=showcol).value) #get the shoe name
        snum = str(ws.cell(row=r,column=epcol).value)   #get the season number
        #print(show,snum)

        #Count the None (empty cells) once we hit MAX_NONE we an assume there are no more shows
        if show == 'None':
            nonecount += 1
            if nonecount == MAX_NONE:
                break
                #Format the dictionary key show:##
        else:
            season = show + ':' + snum.zfill(2)
            #incriment the value by 1, so we cna count number of episodes
            if season in showdict:
                n = showdict[season]
                n += 1
                showdict[season] = n
                #if this is the first time we enountered show:## set the value (show counter to 1)
            else:
                if show != 'Show Title':
                    showdict[season] = 1
        r += 1
        print(".",end="",flush=True)

    return showdict




####Main Program####
xlf,prefix = getXLF()

#Open the xlf / Progress Report
try:
    wb = openpyxl.load_workbook(filename=xlf, read_only=True)
except:
    print("  ~~Could not open", xlf,"\n")
    sys.exit(1)

sheet = selectWS(wb,xlf)
ws = wb[sheet]

channeldict = readXLF(ws)
print(channeldict)



########
#Eliminating writing to the CSV
#outfname = prefix + ".csv"
#outf = open(outfname,"w")

print("")
#outf.write("Show:Season,# Episodes,Notes,Merge Captions / MXF,Jarvis,Upload XL -> S3,Caption -> S3,Caption -> Box Archive,Cleared V1 In Veritone,Status\n")
print("Show:Season:# Episodes:Notes:Merge Captions / MXF:Jarvis:Upload XL -> S3:Caption -> S3:Caption -> Box Archive:Cleared V1 In Veritone:Status")
for s in sorted(channeldict.keys()):
    if s != "Show Title::Season Number:":
        numeps = channeldict[s]
        showParts = s.split(':')
        #season = showParts[0] + ":" + showParts[1] + ":" + str(numeps) + "\n"
        season = showParts[0] + ":" + showParts[1] + ":" + str(numeps)
        print(season)

#outf.close()
wb.close()
