'''
Created on Mar 3, 2017
@author: Siva_Ninala
'''
import os
import datetime
import requests
import urllib
import collections
import io, json
import pandas
import string
from ConfigParser import SafeConfigParser
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import sys

currentDirectory = os.path.dirname(os.path.realpath(__file__))
output_dir = os.path.join(currentDirectory, 'output')
excelFileName = os.path.join(output_dir, 'Jira_Metrics.xlsx')
config_file = os.path.join(currentDirectory, 'config','jiraMetrics.ini')
if os.path.exists(config_file):
    config = SafeConfigParser()
    config.read(config_file)
else:
    raise Exception( config_file + " not found" )

baseUrl = config.get('API', 'search_api_url')
userName = config.get('BUG_TRACKER', 'username')
password = config.get('BUG_TRACKER', 'password')

jsonOutputDir = os.path.join(currentDirectory, 'json')
excelOutputDir = os.path.join(currentDirectory, 'excel')

if not os.path.exists(output_dir):
    os.makedirs(output_dir)
if not os.path.exists(jsonOutputDir):
    os.makedirs(jsonOutputDir)
if not os.path.exists(excelOutputDir):
    os.makedirs(excelOutputDir)

if os.path.exists(excelFileName):
    try:
        os.rename(excelFileName, excelFileName)
    except OSError as e:
        raise OSError(excelFileName + ' already in use. Please close it')

ertProjects = [project.strip() for project in config.get('BUG_TRACKER', 'projects').split(",")]

status_arr = []
jql_items = config.items('JQL')
for status, query in jql_items:
    status_arr.append(status)
status_arr = [string.capwords(status) for status in status_arr]

currentDate = datetime.date.today()
currentWeek = currentDate.strftime("%W-%Y")
currentDate_YYYY_MM_DD = currentDate.strftime("%Y-%m-%d")
currentDate = currentDate.strftime("%m/%d/%Y")

def writeResponseToFileSystem(project, status, response):
    jsonFile = os.path.join(jsonOutputDir, project + '-' + status + '-' + currentDate_YYYY_MM_DD + '.json')
    with io.open(jsonFile, 'w', encoding='utf-8') as f:
        f.write(json.dumps(response, ensure_ascii=False))
    excelFile = os.path.join(excelOutputDir, project + '-' + status + '-' + currentDate_YYYY_MM_DD + '.xlsx')
    pandas.read_json(jsonFile).to_excel(excelFile)


def getResponseFromJira(project, status, query):
    query_string = urllib.quote_plus(query)
    response = requests.get(baseUrl + query_string + '&maxResults=1',
                            auth=HTTPBasicAuth(userName, password))
    if response.status_code == 200:
        response_json = response.json()
        # writeResponseToFileSystem(project, status, response_json)
        return response_json['total']
    else:
        raise Exception("Unable get response from Jira")


def createNewWeeklyMetricsSheet(workBook, workSheet):
    metrics = collections.OrderedDict()
    for project in ertProjects:
        ws = workBook[project]
        for row in ws.iter_rows():
            rowList = []
            for cell in row:
                if cell.row == 1:
                    continue
                rowList.append(cell.value)
            if rowList:
                metrics[project + '#' + rowList[1]] = '#'.join(str(v) for v in rowList[2:])
    keys = metrics.keys()
    rundatesSet = set()
    rundates = list()
    for key in keys:
        project, rundate = key.split('#')
        if rundate not in rundatesSet:
            rundatesSet.add(rundate)
            rundates.append(rundate)
    projectrows = [('Date', 'Total')]
    for rundate in rundates:
        total = 0
        for project in ertProjects:
            (New, diff1, InProgess, diff2, closed, diff3) = metrics[project + '#' + rundate].split('#')
            projectTotal = int(New) + int(InProgess) + int(closed)
            total = total + projectTotal
        projectrows.append((rundate, total))
    rows = projectrows
    for row in rows:
        workSheet.append(row)


def updateWeeklyMetricsSheet(workBook, workSheet):
    metrics = collections.OrderedDict()
    for project in ertProjects:
        ws = workBook[project]
        row = ws.max_row
        latest_row = []
        for col in range(ws.min_column, ws.max_column + 1):
            latest_row.append((ws.cell(row=row, column=col).value))
        metrics[project + '#' + latest_row[1]] = '#'.join(str(v) for v in latest_row[2:])
    project, rundate = metrics.keys()[0].split('#')
    total = 0
    for project in ertProjects:
        (New, diff1, InProgess, diff2, closed, diff3) = metrics[project + '#' + rundate].split('#')
        projectTotal = int(New) + int(InProgess) + int(closed)
        total = (total + projectTotal)
    newRow = (rundate, total)
    for row in workSheet.iter_rows():
        (rundate_cell, total_amount_cell) = (row[0], row[1])
        if rundate_cell.value == currentDate:
            print "updating the data for {}".format(currentDate)
            total_amount_cell.value = newRow[1]
            return
    workSheet.append(newRow)

def createWeeklyTotalBarChart():
    chartName = "WeeklyTotals"
    workBook = load_workbook(excelFileName)
    sheets = workBook.get_sheet_names()
    if chartName not in sheets:
        print "Creating Sheet {}".format(chartName)
        workSheet = workBook.create_sheet(chartName, 0)
        workSheet.sheet_properties.tabColor = "1072BA"
        createNewWeeklyMetricsSheet(workBook, workSheet)
    else:
        print "Updating {} Sheet".format(chartName)
        workSheet = workBook.get_sheet_by_name(chartName)
        updateWeeklyMetricsSheet(workBook, workSheet)
    chart1 = BarChart()
    chart1.style = 10
    chart1.title = "Weekly Total - All Tickets"
    chart1.y_axis.title = 'Total'
    chart1.x_axis.title = 'Run Date'
    data = Reference(workSheet, min_col=2, min_row=1, max_row=workSheet.max_row, max_col=workSheet.max_column)
    cats = Reference(workSheet, min_col=1, min_row=2, max_row=workSheet.max_row)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.series[0].trendline = Trendline()
    chart1.series[0].trendline.trendlineType = 'linear'
    chart1.dataLabels = DataLabelList()
    chart1.dataLabels.showVal = True
    workSheet.add_chart(chart1, "H2")
    workBook.save(filename=excelFileName)


if __name__ == '__main__':
    #### Creare Empty worksheets if the file not exists
    if not os.path.exists(excelFileName):
        statusList = ['Current Week', 'Last Week', 'Difference']
        finalStatusList = ['Current Week', 'Last Week', 'Growth']
        row = '2'
        workBook = Workbook()
        ws = workBook.active
        ###### Rollup Section setup ####
        ws.title = "Rollup"
        #### setting Project ###
        ws.merge_cells('A1:A2')
        ws['A1'] = "Project"
        ws.merge_cells('B1:B2')
        ws['B1'] = "Run Date"
        #### seting New section ###
        ws.merge_cells('C1:E1')
        ws['C1'].value = "New"

        for i, j in zip(range(3), range(ord('C'), ord('E') + 1)):
            ws[chr(j) + row] = statusList[i]
        #### seting New section ###
        ws.merge_cells('F1:H1')
        ws['F1'].value = "In Progress"
        for i, j in zip(range(3), range(ord('F'), ord('H') + 1)):
            ws[chr(j) + row] = statusList[i]
        # setting New section #
        ws.merge_cells('I1:K1')
        ws['I1'].value = "Closed"
        for i, j in zip(range(3), range(ord('I'), ord('K') + 1)):
            ws[chr(j) + row] = statusList[i]
        # setting Total section #
        ws.merge_cells('L1:N1')
        ws['L1'].value = "Total"
        for i in range(ord('L'), ord('N') + 1):
            ws[chr(i) + row] = finalStatusList.pop(0)
        freeze_cell = ws['A2']
        ws.freeze_panes = freeze_cell
        sheet_cols = status_arr
        for index in (1, 3, 5):
            sheet_cols.insert(index, 'diff')
        row = ['Week#', 'Run Date'] + sheet_cols

        for project in ertProjects:
            workSheet = workBook.create_sheet(project)
            workSheet.append(row)
            cell = workSheet['A1']
            workSheet.freeze_panes = cell

        header_font = Font(name='Calibri', size=12, bold=True)
        side = Side(border_style='thin', color="FF000000")
        color_fill = PatternFill("solid", fgColor="87CEEB")
        wrap_alignment = Alignment(wrap_text=True, vertical="top", horizontal='center')
        for sheet in workBook.get_sheet_names():
            ws = workBook[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    border = Border(
                        left=cell.border.left,
                        right=cell.border.right,
                        top=cell.border.top,
                        bottom=cell.border.bottom
                    )
                    border.left = border.right = border.top = border.bottom = side
                    cell.alignment = wrap_alignment
                    cell.border = border
                    cell.font = header_font
                    cell.fill = color_fill

        workBook.save(filename=excelFileName)

    workBook = load_workbook(excelFileName)
    #### populating data for project Sheets ###
    projectWorkSheet = workBook[ertProjects[0]]
    lastRunWeek = projectWorkSheet.cell(row=projectWorkSheet.max_row, column=projectWorkSheet.min_column).value
    lastRunDate = projectWorkSheet.cell(row=projectWorkSheet.max_row, column=projectWorkSheet.min_column + 1).value
    script_executed_for_current_week = (projectWorkSheet.max_row > 1 and (lastRunDate == currentDate))
    print script_executed_for_current_week
    '''
    script_executed_for_current_week = (
        projectWorkSheet.max_row > 1 and (lastRunWeek == currentWeek) or (lastRunDate==currentDate)
    )
    '''
    rollUpSheet = workBook['Rollup']
    rollUpSheet_max_row = rollUpSheet.max_row
    if script_executed_for_current_week:
        rollUpSheet_max_row = (rollUpSheet_max_row - len(ertProjects))

    rollupSheetRows = []
    rollupIndex = 1
    lastWeekResults = dict()
    currentWeekResults = dict()
    for project in ertProjects:
        print "working on Metrics for {} project ".format(project)
        workSheet = workBook[project]
        lastWeekResults = dict()
        currentWeekResults = dict()
        row = []
        projectRowIndex = 1
        if config.has_section(project + '_JQL'):
            jqlQueries = collections.OrderedDict(config.items(project + '_JQL'))
        else:
            jqlQueries = collections.OrderedDict(jql_items)
        for status, query in jqlQueries.iteritems():
            status = string.capwords(status)
            query = query.replace('__PROJECTNAME__', project)
            query = query.replace('__CURRENTDATE__', currentDate_YYYY_MM_DD)
            queryCount = getResponseFromJira(project, status, query)
            # time.sleep(10)
            row.append(queryCount)
            print status, queryCount
            currentWeekResults[project + '-' + status] = queryCount
            project_sheet_max_row = workSheet.max_row
            if script_executed_for_current_week:
                project_sheet_max_row = (project_sheet_max_row - 1)

            if project_sheet_max_row == 1:
                lastRunValue = 0
                diff = 0
            elif status == 'New':
                lastRunValue = workSheet['C' + str(project_sheet_max_row)].value
                diff = "=$C${0}-$C${1}".format(project_sheet_max_row + 1, project_sheet_max_row)
            elif status == 'In Progress':
                lastRunValue = workSheet['E' + str(project_sheet_max_row)].value
                diff = "=$E${0}-$E${1}".format(project_sheet_max_row + 1, project_sheet_max_row)
            elif status == 'Closed':
                lastRunValue = workSheet['G' + str(project_sheet_max_row)].value
                diff = "=$G${0}-$G${1}".format(project_sheet_max_row + 1, project_sheet_max_row)
            row.append(diff)
            lastWeekResults[project + '-' + status] = lastRunValue
        row = [currentWeek, currentDate] + row
        if script_executed_for_current_week:
            index = 0
            for col in range(workSheet.min_column, workSheet.max_column + 1):
                workSheet.cell(row=project_sheet_max_row + 1, column=col, value=row[index])
                index += 1
        else:
            workSheet.append(row)
        if rollUpSheet_max_row == 2:
            currentWeekTotal = currentWeekResults[project + '-New'] +\
                currentWeekResults[project + '-In Progress'] + currentWeekResults[project + '-Closed']
            rollupSheetRows.append([project, currentDate,
                                    currentWeekResults[project + '-New'],
                                    lastWeekResults[project + '-New'],
                                    0,
                                    currentWeekResults[project + '-In Progress'],
                                    lastWeekResults[project + '-In Progress'],
                                    0,
                                    currentWeekResults[project + '-Closed'],
                                    lastWeekResults[project + '-Closed'],
                                    0,
                                    "=$C${0}+$F${0}+$I${0}".format(rollUpSheet_max_row + rollupIndex),
                                    0, 0
                                    ])
        else:
            rollupSheetRows.append([project, currentDate,
                                    currentWeekResults[project + '-New'],
                                    lastWeekResults[project + '-New'],
                                    "=$C${0}-$D${0}".format(rollUpSheet_max_row + rollupIndex),
                                    currentWeekResults[project + '-In Progress'],
                                    lastWeekResults[project + '-In Progress'],
                                    "=$F${0}-$G${0}".format(rollUpSheet_max_row + rollupIndex),
                                    currentWeekResults[project + '-Closed'],
                                    lastWeekResults[project + '-Closed'],
                                    "=$I${0}-$J${0}".format(rollUpSheet_max_row + rollupIndex),
                                    "=$C${0}+$F${0}+$I${0}".format(rollUpSheet_max_row + rollupIndex),
                                    "=$D${0}+$G${0}+$J${0}".format(rollUpSheet_max_row + rollupIndex),
                                    "=$L${0}-$M${0}".format(rollUpSheet_max_row + rollupIndex),
                                    ])
        rollupIndex += 1
    #### populate data for Rollup Sheet ###
    if script_executed_for_current_week:
        index = 1
        for row in rollupSheetRows:
            for col in range(rollUpSheet.min_column, rollUpSheet.max_column + 1):
                rollUpSheet.cell(row=rollUpSheet_max_row + index, column=col, value=rollupSheetRows[index-1][col-1])
            index += 1
    else:
        for row in rollupSheetRows:
            rollUpSheet.append(row)
    workBook.save(filename=excelFileName)
    createWeeklyTotalBarChart()
    print "Task Completed"

