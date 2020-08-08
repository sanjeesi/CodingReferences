import xlrd
from datetime import datetime
import requests

MAX_ROWS = 4

def callApi(data):
    API_ENDPOINT = 'http://'
    res = requests.post(url = API_ENDPOINT, json = data)
    print(res.text)

def readExcel(sheet, row):
    incidentID = sheet.cell_value(row, 0)
    createdDate = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    summary = sheet.cell_value(row, 9)
    updatedDate = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    assignee = sheet.cell_value(row, 4)
    # isActive = True
    objectData = {
        "uniqueTaskId": incidentID,
        "ownerName": assignee,
        "taskDetails": summary,
        "assignedBy": "Hchordiy",
        "priority": "Medium",
        "createdDate": createdDate,
        "updateDate": updatedDate,
        "active": True
    }
    return objectData

loc = (r"Details.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(1)
objectList = []
for index in range(1, MAX_ROWS):
    dataFromExcel = readExcel(sheet, index)
    objectList.append(dataFromExcel)

data = {
    "sourceSystem": "UTS",
    "ownerTaskModelList": objectList
}
print(data)
try:
    callApi(data)
except Exception as e:
    print(e)