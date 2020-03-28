# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'InvoiceAutomation.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!


from JobAutomation.mycaa_main import runProgram
from JobAutomation.mycaa_to_invoice import run_docking_invoices
from PyInvoiceMaster.main import excel_to_pdf
from PyQt5 import QtCore, QtGui, QtWidgets
import sys


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("Invoicing Automation")
        MainWindow.resize(808, 485)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.title = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(48)
        self.title.setFont(font)
        self.title.setObjectName("title")

        self.verticalLayout.addWidget(self.title)
        self.collectStudentsBTN = QtWidgets.QPushButton(self.centralwidget)
        self.collectStudentsBTN.setObjectName("collectStudentsBTN")
        self.verticalLayout.addWidget(self.collectStudentsBTN)
        self.studentsToDockingBTN = QtWidgets.QPushButton(self.centralwidget)
        self.studentsToDockingBTN.setObjectName("studentsToDockingBTN")
        self.verticalLayout.addWidget(self.studentsToDockingBTN)
        self.pdfInvoiceCreatorBTN = QtWidgets.QPushButton(self.centralwidget)
        self.pdfInvoiceCreatorBTN.setObjectName("pdfInvoiceCreatorBTN")
        self.verticalLayout.addWidget(self.pdfInvoiceCreatorBTN)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 808, 22))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuEdit = QtWidgets.QMenu(self.menubar)
        self.menuEdit.setObjectName("menuEdit")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.menuFile.addAction(self.actionExit)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuEdit.menuAction())

        self.collectStudentsBTN.clicked.connect(self.clickedCollectStudents)
        self.studentsToDockingBTN.clicked.connect(self.clickedDockingStudents)
        self.pdfInvoiceCreatorBTN.clicked.connect(self.clickedPDFCreator)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate(
            "Invoicing Automation", "Invoicing Automation"))
        self.title.setText(_translate(
            "Invoicing Automation", "Invoicing Automation"))
        self.collectStudentsBTN.setText(_translate(
            "Invoicing Automation", "Run Students Collection (MYCAA)"))
        self.studentsToDockingBTN.setText(_translate(
            "Invoicing Automation", "Run Students To Docking (MYCAA)"))
        self.pdfInvoiceCreatorBTN.setText(_translate(
            "Invoicing Automation", "Run PDF Invoice Creator (MYCAA)"))
        self.menuFile.setTitle(_translate("Invoicing Automation", "File"))
        self.menuEdit.setTitle(_translate("Invoicing Automation", "Edit"))
        self.actionExit.setText(_translate("Invoicing Automation", "Exit"))

    def clickedCollectStudents(self):
        runProgram()

    def clickedDockingStudents(self):
        run_docking_invoices()

    def clickedPDFCreator(self):
        excel_to_pdf()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
