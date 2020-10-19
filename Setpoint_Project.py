import os
import firebase_admin
from constants import *
from firebase_admin import credentials
from firebase_admin import db
import xlsxwriter
from pathlib import Path

cred = credentials.Certificate('set-point-attender-firebase.json')

firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://set-point-attender.firebaseio.com/'
})

ref = db.reference('/')

""" function for monthly Excel """


def monthly_report_for_client(snapshot, company, month, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{company}_{month}_{year}.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    col = 0
    hour_sum = 0

    worksheet.write(row, col, company)
    row += 1
    worksheet.write(row, col, month)
    worksheet.write(row, col + 1, year)
    row += 1
    col += 1

    for day, workers in snapshot.items():
        col += 1
        if 'Total' not in day:
            worksheet.write(row, col, day)
            worksheet.write(row, col + 1, "עובד")
            worksheet.write(row, col + 2, "שעות")
            row += 1
            col += 1
            for worker in workers:
                worksheet.write(row, col, worker)
                worksheet.write(row, col + 1, workers[worker]['worked'])
                if worker != 'Total':
                    hour_sum += workers[worker]['worked']
                row += 1
            row += 1
            col -= 1
        col -= 1
    worksheet.write(row, col, f'{hour_sum} סך הכל שעות')

    workbook.close()


def monthly_work_to_excel(snapshot, employee, month, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{employee}_{month}_{year}.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    col = 0
    hour_sum = 0

    worksheet.write(row, col, f" עובד/ת: {employee}")
    row += 1
    worksheet.write(row, col, "חודש")
    worksheet.write(row + 1, col, month)
    worksheet.write(row, col + 1, "שנה")
    worksheet.write(row + 1, col + 1, year)

    worksheet.write(row, col + 2, "יום")
    worksheet.write(row, col + 3, "שם לקוח")
    worksheet.write(row, col + 4, "שעות")
    row += 1
    col += 2

    for day, activities in snapshot.items():
        if day == 'status':
            continue
        worksheet.write(row, col, day)
        row += 1
        col += 1
        for activity, report_times in activities.items():
            worksheet.write(row, col, activityToHebrew(activity))
            row += 1
            col += 1
            for time, report in report_times.items():
                worksheet.write(row, col, time)
                col += 1
                worksheet.write(row + 1, col, 'התחלה')
                worksheet.write(row + 1, col + 1, report['start'])
                worksheet.write(row + 2, col, 'סיום')
                worksheet.write(row + 2, col + 1, report['finish'])
                worksheet.write(row + 3, col, 'סך הכל')
                worksheet.write(row + 3, col + 1, report['worked'])
                hour_sum += report['worked']
                col -= 1
                row += 4
            col -= 1
        row += 1
        col -= 1

    worksheet.write(row, 1, f'{hour_sum} סך הכל שעות')

    workbook.close()


def startup_data():
    try:
        request = int(input("לעובדים הקש 0, לקיבוצים 1, פרוייקטים 2, חברות 3, שונות 4\n"))
        if request not in range(0, 5):
            print('invalid request. please try again\n')
            exit(1)

        client = input("worker's name:\n").strip() if request == 0 else input("client's name:\n").strip()

        month = int(input("month:\n").strip())
        year = int(input("year:\n").strip())

        if month not in range(1, 13) or year < 2020:
            print('invalid date. please enter valid time\n')
            exit(1)

        print(f"Report for {client}, at {month}/{year}\n")
        return request, client, month, year
    except ValueError:
        print('invalid input. try again')
        exit(1)


def read_data(client, month, year):
    if int(month) < 10:
        month = '0' + str(month)
    print(f'{client}/{year}-{month}')
    snapshot = ref.child(f'{client}/{year}-{month}').get()
    return snapshot


def append_prefix(request, client):
    if request == 0:
        return "Employees/" + employeesToEnglish[client]
    if request == 1:
        return "Kibbutzim/" + kibbutzimToEnglish[client]
    if request == 2:
        return "Companies/" + companiesToEnglish[client]
    if request == 3:
        return "Projects/" + projectsToEnglish[client]
    else:
        return "Other/" + othersToEnglish[client]


def exists_in_DB(request, client, snapshot):
    if snapshot is None:
        return False
    if request not in range(5):
        return False
    else:
        print(
            [
                client in employeesToEnglish.keys(),
                client in kibbutzimToEnglish.keys(),
                client in companiesToEnglish.keys(),
                client in projectsToEnglish.keys(),
                client in othersToEnglish.keys(),
            ][request]
        )
        return [
            client in employeesToEnglish.keys(),
            client in kibbutzimToEnglish.keys(),
            client in companiesToEnglish.keys(),
            client in projectsToEnglish.keys(),
            client in othersToEnglish.keys(),
        ][request]

#
# def main():
#     request, client, month, year = startup_data()
#
#     snapshot = read_data(append_prefix(request, client), month, year)
#     if not exists_in_DB(request, client, snapshot):
#         print('document does not exists, exiting')
#         exit(1)
#     if request != 0:
#         monthly_report_for_client(snapshot, client, month, year, request)
#     else:
#         monthly_work_to_excel(snapshot, client, month, year)
