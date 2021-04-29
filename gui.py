from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import numpy as np
import pandas as pd
import datetime
from csv import writer
import webbrowser
import os

import ctypes
# Force the taskbar icon
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('LKBrilliant.FinanceTracker.v0.1')

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):

        self.fileName = 'book.csv'

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 268)
        MainWindow.setMinimumSize(QtCore.QSize(500, 0))
        MainWindow.setMaximumSize(QtCore.QSize(750, 400))
        MainWindow.setStyleSheet("background-color: #0f6d15;")

        font = QtGui.QFont()
        font.setFamily("Consolas")

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.btn_export = QtWidgets.QPushButton(self.centralwidget)
        self.btn_export.setMinimumSize(QtCore.QSize(150, 35))
        self.btn_export.setObjectName("btn_export")
        self.btn_export.setFont(font)
        self.btn_export.setStyleSheet(open("style_sheets/btn_export.qss","r").read())
        self.btn_export.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_export.clicked.connect(self.exportPressed)
        self.gridLayout.addWidget(self.btn_export, 1, 2, 1, 1)

        spacerItem = QtWidgets.QSpacerItem(200, 40, QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem, 1, 1, 1, 1)

        self.label_status = QtWidgets.QLabel(self.centralwidget)
        self.label_status.setObjectName("label_status")
        self.label_status.setFont(font)
        self.label_status.setStyleSheet(open("style_sheets/lbl_small_out.qss","r").read())
        self.gridLayout.addWidget(self.label_status, 1, 0, 1, 1)

        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.tabWidget.setFont(font)
        self.tabWidget.currentChanged.connect(self.onTabChange)
        self.tab_1 = QtWidgets.QWidget()
        self.tab_1.setObjectName("tab")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.tab_1)
        self.tabWidget.setStyleSheet(open("style_sheets/tabs.qss","r").read())
        self.tabWidget.tabBar().setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.gridLayout_2.setObjectName("gridLayout_2")

        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem1, 0, 2, 1, 1)

        self.label_onHand = QtWidgets.QLabel(self.tab_1)
        self.label_onHand.setFont(font)
        self.label_onHand.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_onHand.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_onHand.setObjectName("label_onHand")
        self.label_onHand.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_onHand, 5, 2, 1, 1)

        self.label_expense = QtWidgets.QLabel(self.tab_1)
        self.label_expense.setFont(font)
        self.label_expense.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_expense.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_expense.setObjectName("label_expense")
        self.label_expense.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_expense, 3, 2, 1, 1)

        self.label_8 = QtWidgets.QLabel(self.tab_1)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.label_8.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_8, 1, 3, 1, 1)

        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem2, 3, 4, 1, 1)

        self.label_5 = QtWidgets.QLabel(self.tab_1)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.label_5.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_5, 5, 1, 1, 1)

        self.label = QtWidgets.QLabel(self.tab_1)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        self.label.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label, 1, 1, 1, 1)

        self.label_3 = QtWidgets.QLabel(self.tab_1)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.label_3.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_3, 3, 1, 1, 1)

        self.label_10 = QtWidgets.QLabel(self.tab_1)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.label_10.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_10, 5, 3, 1, 1)

        self.label_savings = QtWidgets.QLabel(self.tab_1)
        self.label_savings.setFont(font)
        self.label_savings.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_savings.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_savings.setObjectName("label_savings")
        self.label_savings.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_savings, 1, 2, 1, 1)

        spacerItem3 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_2.addItem(spacerItem3, 6, 2, 1, 1)

        spacerItem4 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_2.addItem(spacerItem4, 3, 0, 1, 1)

        self.label_9 = QtWidgets.QLabel(self.tab_1)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.label_9.setStyleSheet(open("style_sheets/lbl_large.qss","r").read())
        self.gridLayout_2.addWidget(self.label_9, 3, 3, 1, 1)

        spacerItem5 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.gridLayout_2.addItem(spacerItem5, 2, 2, 1, 1)

        spacerItem6 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.gridLayout_2.addItem(spacerItem6, 4, 2, 1, 1)

        self.tabWidget.addTab(self.tab_1, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.tab_2)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.dateEdit = QtWidgets.QDateEdit(self.tab_2)
        self.dateEdit.setMinimumSize(QtCore.QSize(140, 35))
        self.dateEdit.setMaximumSize(QtCore.QSize(100, 16777215))
        self.dateEdit.setFont(font)
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setDate(QtCore.QDate.currentDate())
        self.dateEdit.setObjectName("dateEdit")
        self.dateEdit.setStyleSheet(open("style_sheets/date.qss","r").read())
        self.dateEdit.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.gridLayout_3.addWidget(self.dateEdit, 1, 2, 1, 1)

        self.label_11 = QtWidgets.QLabel(self.tab_2)
        self.label_11.setMinimumSize(QtCore.QSize(0, 35))
        self.label_11.setObjectName("label_11")
        self.label_11.setStyleSheet(open("style_sheets/lbl_small.qss","r").read())
        self.gridLayout_3.addWidget(self.label_11, 1, 1, 1, 1)

        self.label_12 = QtWidgets.QLabel(self.tab_2)
        self.label_12.setMinimumSize(QtCore.QSize(0, 35))
        self.label_12.setObjectName("label_12")
        self.label_12.setFont(font)
        self.label_12.setStyleSheet(open("style_sheets/lbl_small.qss","r").read())
        self.gridLayout_3.addWidget(self.label_12, 3, 1, 1, 1)

        self.label_14 = QtWidgets.QLabel(self.tab_2)
        self.label_14.setMinimumSize(QtCore.QSize(0, 35))
        self.label_14.setObjectName("label_14")
        self.label_14.setFont(font)
        self.label_14.setStyleSheet(open("style_sheets/lbl_small.qss","r").read())
        self.gridLayout_3.addWidget(self.label_14, 5, 1, 1, 1)

        spacerItem7 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.gridLayout_3.addItem(spacerItem7, 4, 1, 1, 1)

        spacerItem8 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem8, 3, 9, 1, 1)

        spacerItem9 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_3.addItem(spacerItem9, 6, 6, 1, 1)

        spacerItem10 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_3.addItem(spacerItem10, 0, 6, 1, 1)

        spacerItem11 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        self.gridLayout_3.addItem(spacerItem11, 2, 1, 1, 1)

        self.comboBox = QtWidgets.QComboBox(self.tab_2)
        self.comboBox.setMinimumSize(QtCore.QSize(0, 35))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setFont(font)
        self.comboBox.setStyleSheet(open("style_sheets/comboBox.qss","r").read())
        self.comboBox.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.comboBox.currentIndexChanged.connect(self.comboChanged)
        self.gridLayout_3.addWidget(self.comboBox, 1, 4, 1, 4)

        spacerItem12 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem12, 3, 0, 1, 1)

        self.btn_save = QtWidgets.QPushButton(self.tab_2)
        self.btn_save.setMinimumSize(QtCore.QSize(100, 35))
        self.btn_save.setObjectName("btn_save")
        self.btn_save.setFont(font)
        self.btn_save.setStyleSheet(open("style_sheets/btn_save.qss","r").read())
        self.btn_save.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_save.clicked.connect(self.savePressed)
        self.gridLayout_3.addWidget(self.btn_save, 5, 6, 1, 2)

        self.lineEdit_remark = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_remark.setMinimumSize(QtCore.QSize(0, 35))
        self.lineEdit_remark.setObjectName("lineEdit_remark")
        self.lineEdit_remark.setFont(font)
        self.lineEdit_remark.setStyleSheet(open("style_sheets/lineEdit_default.qss","r").read())
        self.lineEdit_remark.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.lineEdit_remark.returnPressed.connect(self.returnOnRemark)
        self.gridLayout_3.addWidget(self.lineEdit_remark, 3, 2, 1, 6)

        self.label_13 = QtWidgets.QLabel(self.tab_2)
        self.label_13.setMinimumSize(QtCore.QSize(0, 35))
        self.label_13.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_13.setObjectName("label_13")
        self.label_13.setFont(font)
        self.label_13.setStyleSheet(open("style_sheets/lbl_small.qss","r").read())
        self.gridLayout_3.addWidget(self.label_13, 1, 3, 1, 1)

        self.label_15 = QtWidgets.QLabel(self.tab_2)
        self.label_15.setMinimumSize(QtCore.QSize(0, 35))
        self.label_15.setObjectName("label_15")
        self.label_15.setFont(font)
        self.label_15.setStyleSheet(open("style_sheets/lbl_small.qss","r").read())
        self.gridLayout_3.addWidget(self.label_15, 5, 4, 1, 1)

        self.lineEdit_amount = QtWidgets.QLineEdit(self.tab_2)
        self.lineEdit_amount.setMinimumSize(QtCore.QSize(0, 35))
        self.lineEdit_amount.setFont(font)
        self.lineEdit_amount.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.lineEdit_amount.setObjectName("lineEdit_amount")
        self.lineEdit_amount.setStyleSheet(open("style_sheets/lineEdit_default.qss","r").read())
        self.lineEdit_amount.returnPressed.connect(self.savePressed)
        self.lineEdit_amount.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.onlyInt = QtGui.QIntValidator()
        self.lineEdit_amount.setValidator(self.onlyInt)
        self.gridLayout_3.addWidget(self.lineEdit_amount, 5, 2, 1, 2)

        self.tabWidget.addTab(self.tab_2, "")
        self.gridLayout.addWidget(self.tabWidget, 0, 0, 1, 3)

        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.getBalance()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Finance Tracker v0.1")) 
        MainWindow.setWindowIcon(QtGui.QIcon('UI_graphics/logo_256.png'))        
        self.btn_export.setText(_translate("MainWindow", "Export HTML"))
        self.label_status.setText(_translate("MainWindow", "Status:"))
        self.label_onHand.setText(_translate("MainWindow", ""))
        self.label_expense.setText(_translate("MainWindow", ""))
        self.label_8.setText(_translate("MainWindow", "Rs."))
        self.label_5.setText(_translate("MainWindow", "Can Expend: "))
        self.label.setText(_translate("MainWindow", "Savings: "))
        self.label_3.setText(_translate("MainWindow", "Expenses: "))
        self.label_10.setText(_translate("MainWindow", "Rs."))
        self.label_savings.setText(_translate("MainWindow", ""))
        self.label_9.setText(_translate("MainWindow", "Rs."))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_1), _translate("MainWindow", "Summary"))
        self.label_11.setText(_translate("MainWindow", "Date"))
        self.label_12.setText(_translate("MainWindow", "Remark"))
        self.label_14.setText(_translate("MainWindow", "Amount"))
        self.comboBox.setItemText(0, _translate("MainWindow", "Income"))
        self.comboBox.setItemText(1, _translate("MainWindow", "Expense"))
        self.btn_save.setText(_translate("MainWindow", "Save"))
        self.label_13.setText(_translate("MainWindow", "Type"))
        self.label_15.setText(_translate("MainWindow", "Rs."))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "New Entry"))

    def getBalance(self):
        data = np.genfromtxt('book.csv',delimiter=',', missing_values=',,',skip_header=True)
        # [ID,DATE,Income,Income-Amount,Expense,Expense-Amount]
        colSum = np.nansum(data,axis=0)
        sub = colSum[3] - colSum[5]
        self.label_savings.setText("{:,.0f}".format(sub))
        self.label_expense.setText("{:,.0f}".format(colSum[5]))
        self.label_onHand.setText("{:,.0f}".format(sub*0.2))

    def append_list_as_row(self,list_of_elem,file):
        with open(file, 'a+', newline='') as write_obj:
            csv_writer = writer(write_obj)
            csv_writer.writerow(list_of_elem)

    def onTabChange(self):
        self.getBalance()
        if (self.tabWidget.currentIndex() == 0):
            MainWindow.setStyleSheet("background-color: #44baf4;")
        else:
            mType = self.comboBox.currentText()
            if (mType == 'Income'): MainWindow.setStyleSheet("background-color: #44f490;")
            if (mType == 'Expense'): MainWindow.setStyleSheet("background-color: #f44461;")

    def savePressed(self):
        if self.fieldTest():
            now = datetime.datetime.now()
            timestamp = now.strftime("%d%m%y%H%M%S")

            d = self.dateEdit.date()
            date = '{:02d}-{:02d}-{}'.format(d.day(),d.month(),d.year())

            # [ID,DATE,Income,Income-Amount,Expense,Expense-Amount]
            l = [timestamp, date,'','','','']

            mType = self.comboBox.currentText()
            remark = self.lineEdit_remark.text()
            amount = self.lineEdit_amount.text()

            if (mType == 'Income'):
                l[2] = remark
                l[3] = amount

            elif (mType == 'Expense'):
                l[4] = remark
                l[5] = amount
    
            self.append_list_as_row(l,'book.csv')
            self.lineEdit_remark.setText('')
            self.lineEdit_amount.setText('')
            self.lineEdit_remark.setFocus()

    def returnOnRemark(self):
        self.lineEdit_amount.setFocus()

    def fieldTest(self):
        if (self.lineEdit_remark.text() == ''):
            self.lineEdit_remark.setStyleSheet(open("style_sheets/lineEdit_warning.qss","r").read())
            self.label_status.setText("Status: \'Remark\' field cannot be empty")

        if (self.lineEdit_amount.text() == ''):
            self.lineEdit_amount.setStyleSheet(open("style_sheets/lineEdit_warning.qss","r").read())
            self.label_status.setText("Status: \'Amount\' field cannot be empty")

        if (not(self.lineEdit_amount.text().isdigit())):
            self.lineEdit_amount.setStyleSheet(open("style_sheets/lineEdit_warning.qss","r").read())
            self.label_status.setText("Status: Invalid value on \'Amount\'")

        if (self.lineEdit_remark.text() != '' and self.lineEdit_amount.text() != '' and self.lineEdit_amount.text().isdigit()):
            self.lineEdit_remark.setStyleSheet(open("style_sheets/lineEdit_default.qss","r").read())
            self.lineEdit_amount.setStyleSheet(open("style_sheets/lineEdit_default.qss","r").read())
            return True

    def exportPressed(self):
        df = pd.read_csv('book.csv')
        df.to_html('book.html',index=False,na_rep='',col_space=150,border=0,justify='center')
        webbrowser.open('file://' + os.path.realpath('book.html'))

    def comboChanged(self):
        mType = self.comboBox.currentText()
        if (mType == 'Income'): MainWindow.setStyleSheet("background-color: #44f490;")
        if (mType == 'Expense'): MainWindow.setStyleSheet("background-color: #f44461;")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())