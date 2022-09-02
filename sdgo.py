import json 
import sys
from openpyxl import load_workbook

def createDict():
    wb = load_workbook(filename='sdgo.xlsx')
    sheet = wb['First Sheet']
    unitDict = open("unitDict.json", "w", encoding="UTF-8")
    unit = {}
    for i in range (2, 770):
        a = "A" + str(i)
        d = "D" + str(i)
        unit[sheet[a].value] = sheet[d].value
    unitDict.write(json.dumps(unit, indent=2, ensure_ascii=False))
    unitDict.close()

def intToHexInvert(input):
    hexInput = hex(input)[2:]
    first = hexInput[2:4]
    second = hexInput[0:2]
    return first + second

def HexInvertToint(input):
    first = input[2:4]
    second = input[0:2]
    return str(int(first + second, base = 16))

def changeHex(pos, unitID, accData):
    fileName = sys.argv[1]
    jsonFile = open(fileName, "w", encoding = 'utf-8')
    unitList = accData["grid"]["list"]
    for i in range(0, len(unitList)):
        unitPos = unitList[i]["Pos"]
        if(unitPos == pos):
            unitList[i]["ID"] = intToHexInvert(int(unitID))
            jsonFile.write(json.dumps(accData, ensure_ascii=False)) 
    jsonFile.close()

fileName = sys.argv[1]
jsonFile = open(fileName, "r+", encoding = 'utf-8')
accData = json.load(jsonFile)
unitList = accData["grid"]["list"]
unitDict = open("unitDict.json", "r", encoding = "utf-8")
unitData = json.load(unitDict)
posList = []
for i in range(0, len(unitList)):
    nameIndex = HexInvertToint(unitList[i]["ID"])
    unitPos = unitList[i]["Pos"]
    posList.append(unitPos)
    print(unitPos, unitData[nameIndex])
exit = True
jsonFile.close()
while exit :
    pos = input("请输入要更换的位置或者exit推出 \n")
    if pos == "exit" :
        exit = False
    else :
        if int(pos) in posList:
            unitID = input("输入要换的ID五位数 \n")
            if unitID in unitData.keys():
                changeHex(int(pos), unitID, accData)
            else :
                print("错误的机体ID")
        else :
            print("pos不存在")