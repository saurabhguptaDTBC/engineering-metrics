import json
from openpyxl.styles import Font
import requests
from config import *

def resetSheet(sheetName, workbook):
    if sheetName in workbook.sheetnames:
        del workbook[sheetName]
    workbook.create_sheet(sheetName)
    sheet = workbook[sheetName]
    setSheetHeaderRow(sheet,sheetName)


def setSheetHeaderRow(sheet, sheetName):
    if sheetName == 'Stories':
        sheet.append(['Id','Name','Effort','Project','Team','Feature','LeadTime','CycleTime','Release','Iteration','State','BugsCount','IterationReleaseCount'])
    elif sheetName == 'Releases':
        sheet.append(['Id','Name','EndDate','Total Effort','Release Owner'])
    vHeaderFont=Font(size=14,bold=True)
    for cell in sheet["1:1"]:
        cell.font=vHeaderFont

def ifnull(var, returnVar, subField):
  if var is None:
    return returnVar
  return var[subField]

def delete(sheet):
    while(sheet.max_row > 0):
        sheet.delete(1)
    return

def requestHelper(URL):
    try:
        req = requests.get(URL+gvTPToken)
    except requests.exceptions.RequestException as e:
        print("Error for URL" + URL)
        return "Error"
    textData=req.text
    try:
        jsonDict=json.loads(textData)
    except:
        print("Error with JSON Decoding for "+URL)
        return "Error"
    return jsonDict