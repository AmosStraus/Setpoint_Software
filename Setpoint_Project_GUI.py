import os
import sys

from PyQt5 import QtGui
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import *
from Setpoint_Project import *
from constants import *

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
        self.clientOptions = ('אני', 'אתה')

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
        self.completer2 = QCompleter(allClientsToEnglish)
        self.clientNameDisplay.setCompleter(self.completer2)
        layout.addRow(self.clientNameButton, self.clientNameDisplay)

        self.monthButton = QPushButton("חודש:")
        self.monthButton.clicked.connect(self.getMonth)
        self.monthDisplay = QLineEdit()
        layout.addRow(self.monthButton, self.monthDisplay)

        self.yearButton = QPushButton("שנה:")
        self.yearButton.clicked.connect(self.getYear)
        self.yearDisplay = QLineEdit()
        layout.addRow(self.yearButton, self.yearDisplay)

        self.pathButton = QPushButton("מיקום לשמירה:")
        self.pathButton.clicked.connect(self.getPath)
        self.pathDisplay = QLineEdit()
        layout.addRow(self.pathButton, self.pathDisplay)

        self.finishAndGetButton = QPushButton('קבל דו"ח חודשי')
        self.finishAndGetButton.clicked.connect(self.show_popup)
        layout.addRow(self.finishAndGetButton)

        self.setLayout(layout)
        self.setWindowTitle("סט פוינט הפקת דוחות חודשיים")

    def show_popup(self):
        if self.allFieldsAreValid():
            clientTypeInt = clientTypesToIntMap[self.clientTypeDisplay.text()]
            client = self.clientNameDisplay.text().strip()
            month = self.monthDisplay.text().strip()
            year = self.yearDisplay.text().strip()
            path = self.pathDisplay.text().strip()

            print('before getting data')
            snapshot = read_data(append_prefix(clientTypeInt, client), month, year)
            print(snapshot)
            if exists_in_DB(clientTypeInt, client, snapshot):

                if clientTypeInt != 0:
                    monthly_report_for_client(snapshot, client, month, year, path)
                else:
                    monthly_work_to_excel(snapshot, client, month, year, path)

                msg = QMessageBox()
                msg.setWindowTitle("הושלם בהצלחה")
                msg.setText(
                    f" נוצר דוח אקסל בשם {client}_{month}_{year}.xlsx")
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
        print(i.text())
        if not i.text() == 'OK':
            self.clientTypeDisplay.setText('')
            self.clientNameDisplay.setText('')
            self.monthDisplay.setText('')
            self.yearDisplay.setText('')
            self.pathDisplay.setText('')

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
        num, ok = QInputDialog.getInt(self, "integer input dialog", "הכנס חודש", 0, 1, 12)
        if ok and num is not 0:
            self.monthDisplay.setText(str(num))

    def getYear(self):
        num, ok = QInputDialog.getInt(self, "integer input dialog", "הכנס שנה", 2020, 2020, 2100)
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
        print(self.clientTypeDisplay.text())
        print(self.clientNameDisplay.text())
        print(self.monthDisplay.text())
        print(self.yearDisplay.text())
        if self.pathDisplay.text() is None or not self.pathDisplay.text():
            self.pathDisplay.setText(os.path.abspath(os.getcwd()))
            print(self.pathDisplay.text())

        if (self.clientTypeDisplay.text() in self.clientTypes
                and self.clientNameDisplay.text() in clientTypeStrToEnglish[self.clientTypeDisplay.text()].keys()
                and self.monthDisplay.text().isnumeric()
                and 0 < int(self.monthDisplay.text()) < 13
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
