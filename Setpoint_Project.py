import firebase_admin
from constants import *
from firebase_admin import credentials
from firebase_admin import db
import xlsxwriter

cred = credentials.Certificate('set-point-attender-firebase.json')

firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://set-point-attender.firebaseio.com/'
})

ref = db.reference('/')


def get_added():
    return dict(ref.child('Added_Clients').get())


""" function for monthly Excel """


def yearly_report_for_client(client_type_int, company, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{company}_שנתי_{year}.xlsx')
    for month in range(1, 13):
        snapshot = read_data(append_prefix(client_type_int, company), month, year)

        if exists_in_DB(client_type_int, company, snapshot):
            _monthly_report_for_client(snapshot, company, month, year, workbook)

    workbook.close()


def monthly_report_for_client(snapshot, company, month, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{company}_{month}_{year}.xlsx')
    _monthly_report_for_client(snapshot, company, month, year, workbook)
    workbook.close()


def _monthly_report_for_client(snapshot, company, month, year, workbook):
    worksheet = workbook.add_worksheet(f'{month}_{year}')
    cell_format = workbook.add_format()
    cell_format.set_bg_color('yellow')

    title_format = workbook.add_format()
    title_format.set_bg_color('#FF6E6E')

    first_day_format = workbook.add_format()
    first_day_format.set_bg_color('#55CE55')

    row = 0
    col = 0
    hour_sum = 0

    worksheet.write(row, col, company)
    row += 1

    worksheet.write(row, col, f'דו"ח חודשי')
    worksheet.write(row, col + 1, f"{month}-{year}")
    row += 1

    worksheet.write(row, col, 'תאריך', title_format)
    worksheet.write(row, col + 1, 'עובד/ת', title_format)
    worksheet.write(row, col + 2, 'שעות', title_format)
    worksheet.write(row, col + 3, "סוג פגישה", title_format)
    worksheet.write(row, col + 4, "הערות", title_format)
    worksheet.write(row, col + 5, 'סך השעות היומי', title_format)
    row += 1

    for day, workers in snapshot.items():
        daily_hour_sum = 0
        temp_row = row
        for worker in workers:
            worksheet.write(row, col, day, first_day_format) if temp_row == row else worksheet.write(row, col, day)
            worksheet.write(row, col + 1, employeesToHebrew[worker])
            worksheet.write(row, col + 2, workers[worker]['worked'])
            meetingType = meetingTypeMap[worker['meetingType']] if 'meetingType' in worker else "אחר"
            worksheet.write(row, col+3, meetingType)
            comment = worker['comment'] if 'comment' in worker else ""
            worksheet.write(row, col+4, comment)
            row += 1
            daily_hour_sum += workers[worker]['worked']

        worksheet.write(temp_row, col + 5, daily_hour_sum, first_day_format)
        hour_sum += daily_hour_sum

    worksheet.write(row, col + 2, 'סך שעות חודשי', cell_format)
    worksheet.write(row, col + 3, hour_sum, cell_format)


def monthly_employee_work_to_excel(snapshot, employee, month, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{employee}_{month}_{year}.xlsx')
    _monthly_employee_work_to_excel(snapshot, employee, month, year, workbook)
    workbook.close()


def yearly_employee_work_to_excel(client_type_int, employee, year, path):
    workbook = xlsxwriter.Workbook(f'{path}\\{employee}_שנתי_{year}.xlsx')
    for month in range(1, 13):
        snapshot = read_data(append_prefix(client_type_int, employee), month, year)

        if exists_in_DB(client_type_int, employee, snapshot):
            _monthly_employee_work_to_excel(snapshot, employee, month, year, workbook)

    workbook.close()


def _monthly_employee_work_to_excel(snapshot, employee, month, year, workbook):
    worksheet = workbook.add_worksheet(f'{month}_{year}')
    cell_format = workbook.add_format()
    cell_format.set_bg_color('yellow')

    title_format = workbook.add_format()
    title_format.set_bg_color('#FF6E6E')

    first_day_format = workbook.add_format()
    first_day_format.set_bg_color('#55CE55')

    row = 0
    col = 0
    hour_sum = 0

    worksheet.write(row, col, 'שם העובד/ת')
    worksheet.write(row, col + 1, employee)

    row += 1

    worksheet.write(row, col, f'דו"ח חודשי')
    worksheet.write(row, col + 1, f"{month}-{year}")

    row += 1

    worksheet.write(row, col, "שם עובד", title_format)
    worksheet.write(row, col + 1, "תאריך", title_format)
    worksheet.write(row, col + 2, "מהות", title_format)
    worksheet.write(row, col + 3, "סך הכל", title_format)
    worksheet.write(row, col + 4, "סוג פגישה", title_format)
    worksheet.write(row, col + 5, "הערות", title_format)
    worksheet.write(row, col + 6, "סך שעות יומי", title_format)
    row += 1

    for day, activities in snapshot.items():
        daily_sum = 0
        if day == 'status':
            continue
        temp_row = row
        for activity, report_times in activities.items():
            for time, report in report_times.items():
                worksheet.write(row, col, employee)  # worker name
                if temp_row == row:
                    worksheet.write(row, col + 1, day, first_day_format)
                else:
                    worksheet.write(row, col + 1, day)  # date
                worksheet.write(row, col + 2, activityToHebrew(activity))  # essence
                worksheet.write(row, col + 3, report['worked'])
                meetingType = meetingTypeMap[report['meetingType']] if 'meetingType' in report else "אחר"
                worksheet.write(row, col + 4, meetingType)
                comment = report['comment'] if 'comment' in report else ""
                worksheet.write(row, col + 5, comment)

                hour_sum += report['worked']
                daily_sum += report['worked']
                row += 1

        worksheet.write(temp_row, col + 6, daily_sum, first_day_format)
        row += 1

    worksheet.write(row, 2, 'סך שעות חודשי', cell_format)
    worksheet.write(row, 3, hour_sum, cell_format)


def read_data(client, month, year):
    if int(month) < 10:
        month = '0' + str(month)

    snapshot = ref.child(f'{client}/{year}-{month}').get()
    return snapshot


def append_prefix(request, client):
    if request == 0:
        return "Employees/" + (employeesToEnglish[client] if client in employeesToEnglish else client)
    if request == 1:
        return "Kibbutzim/" + (kibbutzimToEnglish[client] if client in kibbutzimToEnglish else client)
    if request == 2:
        return "Companies/" + (companiesToEnglish[client] if client in companiesToEnglish else client)
    if request == 3:
        return "Projects/" + (projectsToEnglish[client] if client in projectsToEnglish else client)
    else:
        return "Other/" + (othersToEnglish[client] if client in othersToEnglish else client)


def exists_in_DB(request, client, snapshot):
    if snapshot is None:
        return False
    if request not in range(5):
        return False
    if client in get_added():
        return True
    else:
        return [
            client in employeesToEnglish.keys(),
            client in kibbutzimToEnglish.keys(),
            client in companiesToEnglish.keys(),
            client in projectsToEnglish.keys(),
            client in othersToEnglish.keys(),
        ][request]


def activityToHebrew(activity):
    if activity in companiesToHebrew.keys():
        return companiesToHebrew[activity]

    elif activity in kibbutzimToHebrew.keys():
        return kibbutzimToHebrew[activity]

    elif activity in othersToHebrew.keys():
        return othersToHebrew[activity]

    elif activity in projectsToHebrew.keys():
        return projectsToHebrew[activity]

    elif activity in employeesToHebrew.keys():
        return employeesToHebrew[activity]
    # elif activity in get_added().keys():
    #     return get_added()[activity]
    else:
        return activity
