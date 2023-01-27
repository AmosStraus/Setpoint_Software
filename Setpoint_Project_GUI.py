import os
import sys
# auto-py-to-exe
from PyQt5 import QtGui
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import *
from Setpoint_Project import *
from constants import *
from datetime import datetime

clientTypesToIntMap = {
    'עובדים': 0,
    'קיבוצים': 1,
    'חברות': 2,
    'פרוייקטים': 3,
    'אחר': 4
}


class inputUI(QWidget):
    def __init__(self, parent=None):
        super(inputUI, self).__init__(parent)

        self.setGeometry(400, 250, 800, 600)
        self.clientTypes = ('עובדים', 'קיבוצים', 'חברות', 'פרוייקטים', 'אחר')
        self.clientOptions = ('','קודם בחר/י קטגוריה')

        layout = QFormLayout()
        # left, top. right. bottom.
        layout.setContentsMargins(60, 60, 60, 60)

        buttonFont = QFont()
        buttonFont.setPointSize(16)
        inputFont = QFont()
        inputFont.setPointSize(18)
        QApplication.setFont(buttonFont, "QLabel")
        QApplication.setFont(buttonFont, "QPushButton")
        QApplication.setFont(inputFont, "QInputDialog")
        QApplication.setFont(inputFont, "QLineEdit")

        logo = QLabel()
        logo.setPixmap(QPixmap("assets\\setpoint_logo_small.jpg"))
        header = QLabel()
        header.setText('פרטים לדיווח חודשי\n(למלא בעברית) \n לחץ על הכפתור משמאל לאפשרויות')

        # must do it that way
        myFont = QtGui.QFont('Times', 18)
        myFont.setBold(True)
        myFont.setUnderline(True)
        header.setFont(myFont)
        layout.addRow(logo, header)

        self.clientTypeButton = QPushButton("בחר קטגוריה:")
        self.clientTypeButton.clicked.connect(self.getClientType)
        self.clientTypeDisplay = QLineEdit()
        completer1 = QCompleter(self.clientTypes)
        self.clientTypeDisplay.setCompleter(completer1)
        layout.addRow(self.clientTypeButton, self.clientTypeDisplay)

        self.clientNameButton = QPushButton("שם עובד/לקוח:")
        self.clientNameButton.clicked.connect(self.getClientName)
        self.clientNameDisplay = QLineEdit()
        self.completer2 = QCompleter({**allClientsToEnglish, **get_added()})
        self.clientNameDisplay.setCompleter(self.completer2)
        layout.addRow(self.clientNameButton, self.clientNameDisplay)

        self.monthButton = QPushButton("חודש:")
        self.monthButton.clicked.connect(self.getMonth)
        self.monthDisplay = QLineEdit(str(datetime.now().month))
        layout.addRow(self.monthButton, self.monthDisplay)

        self.yearButton = QPushButton("שנה:")
        self.yearButton.clicked.connect(self.getYear)
        self.yearDisplay = QLineEdit(str(datetime.now().year))
        layout.addRow(self.yearButton, self.yearDisplay)

        self.pathButton = QPushButton("מיקום לשמירה:")
        self.pathButton.clicked.connect(self.getPath)
        self.pathDisplay = QLineEdit()
        layout.addRow(self.pathButton, self.pathDisplay)

        self.finishAndGetButton = QPushButton('קבל דו"ח חודשי')
        self.finishAndGetButton.clicked.connect(self.show_popup_monthly)
        layout.addRow(self.finishAndGetButton)

        self.finishAndGetButtonYear = QPushButton('קבל דו"ח שנתי')
        self.finishAndGetButtonYear.clicked.connect(self.show_popup_yearly)
        layout.addRow(self.finishAndGetButtonYear)

        self.setLayout(layout)
        self.setWindowTitle("סט פוינט הפקת דוחות חודשיים")

    def show_popup_monthly(self):
        if self.allFieldsAreValid():
            clientTypeInt = clientTypesToIntMap[self.clientTypeDisplay.text()]
            client = self.clientNameDisplay.text().strip()
            month = self.monthDisplay.text().strip()
            year = self.yearDisplay.text().strip()
            path = self.pathDisplay.text().strip()
            snapshot = read_data(append_prefix(clientTypeInt, client), month, year)

            if snapshot and exists_in_DB(clientTypeInt, client, snapshot):

                if clientTypeInt != 0:
                    monthly_report_for_client(snapshot, client, month, year, path)
                else:
                    monthly_employee_work_to_excel(snapshot, client, month, year, path)

                msg = QMessageBox()
                msg.setWindowTitle("הושלם בהצלחה")
                msg.setText(
                    f" נוצר דוח חודשי בשם {client}_{month}_{year}.xlsx")
                msg.setIcon(QMessageBox.Question)
                msg.addButton(QMessageBox.Ok)
                msg.addButton('להוצאת דו"ח אחר', QMessageBox.YesRole)
                msg.setInformativeText(f" בנתיב{self.pathDisplay.text()}")
                msg.buttonClicked.connect(self.popup_button)
                msg.exec_()
            else:
                errorMsg = QMessageBox()
                errorMsg.setWindowTitle("מסמך לא קיים")
                errorMsg.setText("אין דיווח במערכת עבור הנתונים שהוזנו")
                errorMsg.setIcon(QMessageBox.Critical)
                errorMsg.exec_()
        else:
            errorMsg = QMessageBox()
            errorMsg.setWindowTitle("פרטים אינם תקינים")
            errorMsg.setText("מלא/י את כל השדות הנחוצים באופן חוקי")
            errorMsg.setIcon(QMessageBox.Critical)
            errorMsg.exec_()

    def popup_button(self, i):
        if not i.text() == 'OK':
            self.clientTypeDisplay.setText('')
            self.clientNameDisplay.setText('')
            self.monthDisplay.setText('')
            self.yearDisplay.setText('')
            self.pathDisplay.setText('')

    def show_popup_yearly(self):
        if self.allFieldsAreValidYear():
            clientTypeInt = clientTypesToIntMap[self.clientTypeDisplay.text()]
            client = self.clientNameDisplay.text().strip()
            year = self.yearDisplay.text().strip()
            path = self.pathDisplay.text().strip()

            if clientTypeInt != 0:
                yearly_report_for_client(clientTypeInt, client, year, path)
            else:
                yearly_employee_work_to_excel(clientTypeInt, client, year, path)

            msg = QMessageBox()
            msg.setWindowTitle("הושלם בהצלחה")
            msg.setText(
                f" נוצר דוח שנתי בשם {client}_שנתי_{year}.xlsx")
            msg.setIcon(QMessageBox.Question)
            msg.addButton(QMessageBox.Ok)
            msg.addButton('להוצאת דו"ח אחר', QMessageBox.YesRole)
            msg.setInformativeText(f" בנתיב{self.pathDisplay.text()}")
            # msg.buttonClicked.connect(self.popup_button)
            msg.exec_()
        else:
            errorMsg = QMessageBox()
            errorMsg.setWindowTitle("פרטים אינם תקינים")
            errorMsg.setText("מלא/י את כל השדות הנחוצים באופן חוקי")
            errorMsg.setIcon(QMessageBox.Critical)
            errorMsg.exec_()

    def getClientType(self):
        item, ok = QInputDialog.getItem(self, "select input dialog",
                                        "רשימת קטגוריות", self.clientTypes, 0, False)
        if ok and item:
            self.clientTypeDisplay.setText(str(item))
            self.clientOptions = clientTypeStrToEnglish[str(item).strip()]
            self.completer2 = QCompleter(clientTypeStrToEnglish[str(item).strip()])
            self.clientNameDisplay.setCompleter(self.completer2)

    def getClientName(self):
        item, ok = QInputDialog.getItem(self, "select input dialog",
                                        "רשימת קטגוריות", self.clientOptions, 0, False)
        if ok and item != '':
            self.clientNameDisplay.setText(str(item))

    def getMonth(self):
        num, ok = QInputDialog.getInt(self, "integer input dialog", "הכנס חודש", datetime.now().month, 1, 12)
        if ok and num is not 0:
            self.monthDisplay.setText(str(num))

    def getYear(self):
        num, ok = QInputDialog.getInt(self, "integer input dialog", "הכנס שנה", datetime.now().year, 2020, 2100)
        if ok and num is not 0:
            self.yearDisplay.setText(str(num))

    def getPath(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        options |= QFileDialog.DontUseCustomDirectoryIcons
        dialog = QFileDialog()
        dialog.setOptions(options)

        path = dialog.getExistingDirectory(self, 'בחר מקום לשמור בו את הדו"ח')
        self.pathDisplay.setText(path)

    def allFieldsAreValid(self):
        if self.pathDisplay.text() is None or not self.pathDisplay.text():
            self.pathDisplay.setText(os.path.abspath(os.getcwd()))

        if (self.clientTypeDisplay.text() in self.clientTypes
                # and self.clientNameDisplay.text() in clientTypeStrToEnglish[self.clientTypeDisplay.text()].keys()
                and self.monthDisplay.text().isnumeric()
                and 0 < int(self.monthDisplay.text()) < 13
                and self.yearDisplay.text().isnumeric()
                and 2019 < int(self.yearDisplay.text())):
            return True
        return False

    def allFieldsAreValidYear(self):
        if self.pathDisplay.text() is None or not self.pathDisplay.text():
            self.pathDisplay.setText(os.path.abspath(os.getcwd()))

        if (self.clientTypeDisplay.text() in self.clientTypes
                # and self.clientNameDisplay.text() in clientTypeStrToEnglish[self.clientTypeDisplay.text()].keys()
                and self.yearDisplay.text().isnumeric()
                and 2019 < int(self.yearDisplay.text())):
            return True
        return False


def main():
    app = QApplication(sys.argv)
    ex = inputUI()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
