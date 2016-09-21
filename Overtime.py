from __future__ import division
from collections import deque
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import sys
import requests
import json
import re
import datetime

excel_filename = 'rc_pm.xlsx'

harvest_headers = {
    'Content-type': 'application/json',
    'Accept': 'application/json',
    'Authorization': 'Basic Y21vcmlrdW5pQHJldmFjb21tLmNvbToqNjAlaEZ4ViVSWHU='
}

def init():
    # Harvest - Build request & load projects
    projects = requests.get('https://revacomm.harvestapp.com/projects', headers=harvest_headers)
    projects_json = projects.json()

    return projects_json


def harvestBudget(project_code, projects):
    # Find project ID
    project_id = None
    project_budget = 0
    starts_on = '20000101'
    today = datetime.datetime.now().strftime("%Y%m%d")
    for project in projects:
        project = project['project']
        if project["code"] == project_code:
            project_id = project["id"]
            project_budget = project["budget"]
            if not project_budget:
                return ("ERROR: No budget has been assigned to " + project["name"] + ".", None, None)
            break

    # BILLABLE
    entries = requests.get('https://revacomm.harvestapp.com/projects/' + str(project_id) + '/entries?from=' + starts_on + '&to=' + today + '&billable=yes', headers=harvest_headers)
    entries_json = entries.json()
    if type(entries_json) is not list and entries_json.get("error", None):
        return ("ERROR: " + project_code + " not found.", None, None)

    billable = 0
    for entry in entries_json:
        entry = entry['day_entry']
        billable += entry['hours']

    burn = billable / project_budget * 100
    remain = 100 - burn
    return (None, burn, remain)


def openExcel(filename):
    wb = None
    ws = None

    # Open or create a new worksheet
    today = datetime.datetime.now().strftime("%Y.%m.%d")
    if os.path.exists(filename):
        wb = load_workbook(filename)
        if today in wb.get_sheet_names():
            ws = wb[today]
    else:
        wb = Workbook()
        ws = wb.active

    # Delete current sheet & create new
    if ws is not None:
        wb.remove_sheet(ws)
    ws = wb.create_sheet(0)
    ws.title = today

    # CM setup sheet headers
    headers = ["Harvest Code", "Wrike Name", "Completion", "Burn", "Remain"]
    for col, header in enumerate(headers):
        c = ws.cell(row = 1, column = col+1)
        c.value = header
    return (wb, ws)


def closeExcel(wb, filename):
    wb.save(filename)


def outputToExcel(ws, project, index):
    proj_tmp = [project["Harvest_Code"], project["Wrike_Name"], project["Progress"]["Completion"], project["Progress"]["Burn"], project["Progress"]["Remain"]]
    for col, val in enumerate(proj_tmp):
        c = ws.cell(row = index, column = col+1)
        c.value = val


def userTotalTime(userTimeJson):
    hours = 0
    timeByDay = {}
    for timeEntry in userTimeJson:
        timeEntry = timeEntry["day_entry"]
        spentAt = timeEntry["spent_at"]

        isWeekend = False
        date = datetime.datetime.strptime(spentAt, '%Y-%m-%d')
        if date.weekday() >= 5:
            isWeekend = True

        if spentAt not in timeByDay:
            timeByDay[spentAt] = {
                "weekend": isWeekend,
                "hours": 0
            }
        timeByDay[spentAt]["hours"] += timeEntry["hours"]
        hours = hours + timeEntry["hours"]

    over = 0
    for date in timeByDay:
        weekend = timeByDay[date]["weekend"]
        dailyHours = timeByDay[date]["hours"]
        if weekend:
            over += dailyHours
        elif dailyHours > 8:
            over += dailyHours - 8
    return (hours, over)

if __name__ == '__main__':
    # dayOfYear = datetime.datetime.now().timetuple().tm_yday + 1
    # for day in range(1, dayOfYear):
    #     print day
    #     entry = requests.get('https://revacomm.harvestapp.com/daily/' + str(day) + '/2016', headers=harvest_headers)
    #     entry_json = entry.json()
    #     print entry_json

    firstDayOfYear = '20160101'
    today = str(datetime.datetime.today().strftime('%Y%m%d'))
    peopleTime = {}
    people = requests.get('https://revacomm.harvestapp.com/people', headers=harvest_headers)
    people_json = people.json()
    for person in people_json:
        pUser = person['user']
        uid = str(pUser['id'])
        first = pUser['first_name']
        last = pUser['last_name']

        userTime = requests.get('https://revacomm.harvestapp.com/people/' + uid + '/entries?from=' + firstDayOfYear + '&to=' + today, headers=harvest_headers)
        userTime_json = userTime.json()
        if not userTime_json:
            continue
        hours, over = userTotalTime(userTime_json)
        print first + ": TOT: " + str(hours) + " OVER: " + str(over)

        peopleTime[uid] = {
            "first": first,
            "last": last,
            "total_hours": hours,
            "overtime": over
        }


    # https://YOURACCOUNT.harvestapp.com/people/{USER_ID}/entries?from=YYYYMMDD&to=YYYYMMDD
    entry = requests.get('https://revacomm.harvestapp.com/people/', headers=harvest_headers)
    entry_json = entry.json()
    #print json.dumps(entry_json)

    emps = [
        ""
    ]

    # Initialize API calls
    #harvest = init()

    # # Loop Codes and Titles
    # index = 2   # Start from 2, headers are in row 1
    # wb, ws = openExcel(excel_filename)
    # for project in input_json["Projects"]:
    #     project_code = project["Harvest_Code"]
    #     project_title = project["Wrike_Name"]
    #
    #     comp_per = wrikeCompletion(project_title, wrike)
    #     budget = harvestBudget(project_code, harvest)
    #     if comp_per[0] is not None:
    #         print comp_per[0]
    #         continue
    #
    #     if budget[0] is not None:
    #         print budget[0]
    #         continue
    #
    #     # Store results in JSON
    #     progress = {
    #         "Completion": comp_per[1],
    #         "Burn": budget[1],
    #         "Remain": budget[2]
    #     }
    #     project["Progress"] = progress
    #
    #     print "\n" + project_title
    #     print "Project Completion: " + str(comp_per[1])
    #     print "Budget Burn: " + str(budget[1])
    #     print "Budget Remaining: " + str(budget[2])
    #
    #     outputToExcel(ws, project, index)
    #     index += 1
    # closeExcel(wb, excel_filename)
