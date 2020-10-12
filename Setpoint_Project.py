from firebase import firebase
import xlsxwriter


""" function for monthly Excel """

def monthlyWorkToExcel(result, company, month, year):
    workbook = xlsxwriter.Workbook(f'{company}_{month}_{year}.xlsx')
    worksheet = workbook.add_worksheet()
    
    row = 0
    col = 0
    
    sum = 0
    indent = 4
    
    
    worksheet.write(row, col, company)
    row += 1
    worksheet.write(row, col, month)
    worksheet.write(row, col + 1, year)
    row += 1
    col += 1
    
    for day, workers in result.items(): 
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
                    sum += workers[worker]['worked']
                row += 1
            row += 1
            col -= 1
        col -= 1
    worksheet.write(row, col, f'{sum} סך הכל שעות')
    
    workbook.close()
  



""" function for monthly Prints """

def printMonthlyWork(result, month):
    indent = 0
    sum = 0
    print(month)
    for day, workers in result.items(): 
        indent += 4
        if 'Total' not in day:
            print("\n" , ' '*indent , day)
            indent += 4
            for worker in workers:
                print(' '*indent, worker + ":".ljust(10 - len(worker)), workers[worker]['worked'], "hours") 
                print(worker)
                if worker != 'Total':
                    sum += workers[worker]['worked']
            print("")
            indent -= 4
        else:
            print("\nMonthly Total", workers, "\n")
        indent -= 4
    print(f'{sum} hours total in {month}\n')
    

fb_app = firebase.FirebaseApplication('https://choosing-names-firebasecodelab.firebaseio.com/')

septemberHatzerim = fb_app.get('Kibbutzim/Hatzerim/2020-09', None)
octoberHatzerim = fb_app.get('Kibbutzim/Hatzerim/2020-10', None)
octoberOrHaner = fb_app.get('Kibbutzim/OrHaner/2020-10', None)

##printMonthlyWork(september, "SEPTEMBER")
##printMonthlyWork(october, "OCTOBER")

monthlyWorkToExcel(septemberHatzerim, "Hatzerim", "SEPTEMBER", '2020')
monthlyWorkToExcel(octoberHatzerim, "Hatzerim", "OCTOBER", '2020')
monthlyWorkToExcel(octoberOrHaner, "OrHaner", "OCTOBER", '2020')