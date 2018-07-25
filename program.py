import os
import random
import re
import xlwt


dataDir = 'test'
NUMBER_OF_CLASS = 5


def numberSplit(txt):
    res = re.findall(r"X:(.*) Y:(.*) P:(.*)", txt)
    X = res[0][0]
    Y = res[0][1]
    Z = res[0][2]
    result = X + " " + Y + " " + Z + '\n'
    return result


def makeExl(filename):
    fopen = open(dataDir+"/"+filename)
    try:
        lines = fopen.readlines()
    except:
        print(filename+"有问题")
    else:
        random.shuffle(lines)

        a = 0
        b = 0
        c = 0
        A = ""
        B = ""
        C = ""
        for line in lines:
            if line[4] == '0' and a < NUMBER_OF_CLASS:
                A += numberSplit(line)
                a += 1

            elif line[4] == '1' and b < NUMBER_OF_CLASS:
                B += numberSplit(line)
                b += 1

            elif line[4] == '2' and c < NUMBER_OF_CLASS:
                C += numberSplit(line)
                c += 1

        Date = "0: \n"+A+'1: \n'+B+"2: \n"+C

        temp = open("temp", 'w')
        temp.write(Date)
        temp.close()

        tempOpen = open("temp")
        tempLines = tempOpen.readlines()

        sheet = exlFile.add_sheet(filename)

        exlX = 0
        exlY = 0
        for tLine in tempLines:
            if ":" in tLine:
                continue
            lineDate = tLine.split(" ")
            sheet.write(exlX, exlY, lineDate[0])
            exlY += 1
            sheet.write(exlX, exlY, lineDate[1])
            exlY += 1
            sheet.write(exlX, exlY, lineDate[2])
            exlX += 1
            exlY = 0

        exlFile.save('RES.xls')
        print(filename + '完成')


exlFile = xlwt.Workbook(encoding='utf-8', style_compression=0)
for files in os.listdir(dataDir):
    makeExl(files)
