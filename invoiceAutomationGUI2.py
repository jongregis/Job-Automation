

from JobAutomation.mycaa_main import runProgram
from JobAutomation.elearning_main import run_program_elearning
from JobAutomation.privatePay_main import run_program_privatePay
from JobAutomation.mycaa_to_invoice import run_docking_invoices
from JobAutomation.elearning_to_invoice import run_docking_invoices_elearning
from JobAutomation.privatePay_to_invoice import run_docking_invoices_privatePay
from PyInvoiceMaster.main import excel_to_pdf, excel_to_pdf_ELearning, excel_to_pdf_PrivatePay
from PyQt5 import QtCore, QtGui, QtWidgets
import sys


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(808, 485)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.MYCAAcollectStudentsBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.MYCAAcollectStudentsBTN.setObjectName("MYCAAcollectStudentsBTN")
        self.verticalLayout.addWidget(self.MYCAAcollectStudentsBTN)
        self.MYCAAstudentsToDockingBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.MYCAAstudentsToDockingBTN.setObjectName(
            "MYCAAstudentsToDockingBTN")
        self.verticalLayout.addWidget(self.MYCAAstudentsToDockingBTN)
        self.MYCAApdfInvoiceCreatorBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.MYCAApdfInvoiceCreatorBTN.setObjectName(
            "MYCAApdfInvoiceCreatorBTN")
        self.verticalLayout.addWidget(self.MYCAApdfInvoiceCreatorBTN)
        self.title = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(48)
        self.title.setFont(font)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setObjectName("title")
        self.verticalLayout.addWidget(self.title)
        self.ElearningCollectStudentsBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.ElearningCollectStudentsBTN.setObjectName(
            "ElearningCollectStudentsBTN")
        self.verticalLayout.addWidget(self.ElearningCollectStudentsBTN)
        self.ElearningstudentsToDockingBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.ElearningstudentsToDockingBTN.setObjectName(
            "ElearningstudentsToDockingBTN")
        self.verticalLayout.addWidget(self.ElearningstudentsToDockingBTN)
        self.ElearningpdfInvoiceCreatorBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.ElearningpdfInvoiceCreatorBTN.setObjectName(
            "ElearningpdfInvoiceCreatorBTN")
        self.verticalLayout.addWidget(self.ElearningpdfInvoiceCreatorBTN)

        # PrivatePay
        self.title.setObjectName("title")
        self.verticalLayout.addWidget(self.title)
        self.PrivatePayCollectStudentsBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.PrivatePayCollectStudentsBTN.setObjectName(
            "PrivatePayCollectStudentsBTN")
        self.verticalLayout.addWidget(self.PrivatePayCollectStudentsBTN)
        self.PrivatePaystudentsToDockingBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.PrivatePaystudentsToDockingBTN.setObjectName(
            "PrivatePaystudentsToDockingBTN")
        self.verticalLayout.addWidget(self.PrivatePaystudentsToDockingBTN)
        self.PrivatePaypdfInvoiceCreatorBTN = QtWidgets.QPushButton(
            self.centralwidget)
        self.PrivatePaypdfInvoiceCreatorBTN.setObjectName(
            "PrivatePaypdfInvoiceCreatorBTN")
        self.verticalLayout.addWidget(self.PrivatePaypdfInvoiceCreatorBTN)

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

        self.MYCAAcollectStudentsBTN.clicked.connect(
            self.clickedCollectStudentsMYCAA)
        self.MYCAAstudentsToDockingBTN.clicked.connect(
            self.clickedDockingStudentsMYCAA)
        self.MYCAApdfInvoiceCreatorBTN.clicked.connect(
            self.clickedPDFCreatorMYCAA)
        # E-learning
        self.ElearningCollectStudentsBTN.clicked.connect(
            self.clickedCollectStudentsElearning)
        self.ElearningstudentsToDockingBTN.clicked.connect(
            self.clickedDockingStudentsElearning)
        self.ElearningpdfInvoiceCreatorBTN.clicked.connect(
            self.clickedPDFCreatorElearning)
        # PrivatePay
        self.PrivatePayCollectStudentsBTN.clicked.connect(
            self.clickedCollectStudentsPrivatePay)
        self.PrivatePaystudentsToDockingBTN.clicked.connect(
            self.clickedDockingStudentsPrivatePay)
        self.PrivatePaypdfInvoiceCreatorBTN.clicked.connect(
            self.clickedPDFCreatorPrivatePay)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.MYCAAcollectStudentsBTN.setText(_translate(
            "MainWindow", "Run Students Collection (MYCAA)"))
        self.MYCAAstudentsToDockingBTN.setText(_translate(
            "MainWindow", "Run Students To Docking (MYCAA)"))
        self.MYCAApdfInvoiceCreatorBTN.setText(_translate(
            "MainWindow", "Run PDF Invoice Creator (MYCAA)"))
        self.title.setText(_translate("MainWindow", "Invoicing Automation"))
        self.ElearningCollectStudentsBTN.setText(_translate(
            "MainWindow", "Run Students Collection (E-Learning)"))
        self.ElearningstudentsToDockingBTN.setText(_translate(
            "MainWindow", "Run Students To Docking (E-Learning)"))
        self.ElearningpdfInvoiceCreatorBTN.setText(_translate(
            "MainWindow", "Run PDF Invoice Creator (E-Learning)"))
        # PrivatePay
        self.PrivatePayCollectStudentsBTN.setText(_translate(
            "MainWindow", "Run Students Collection (PrivatePay)"))
        self.PrivatePaystudentsToDockingBTN.setText(_translate(
            "MainWindow", "Run Students To Docking (PrivatePay)"))
        self.PrivatePaypdfInvoiceCreatorBTN.setText(_translate(
            "MainWindow", "Run PDF Invoice Creator (PrivatePay)"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuEdit.setTitle(_translate("MainWindow", "Edit"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))

    def clickedCollectStudentsMYCAA(self):
        runProgram('12', 'Dec')

    def clickedDockingStudentsMYCAA(self):
        run_docking_invoices()

    def clickedPDFCreatorMYCAA(self):
        excel_to_pdf()
# Elearning

    def clickedCollectStudentsElearning(self):
        run_program_elearning('12')

    def clickedDockingStudentsElearning(self):
        run_docking_invoices_elearning()

    def clickedPDFCreatorElearning(self):
        excel_to_pdf_ELearning()

# Private Pay
    def clickedCollectStudentsPrivatePay(self):
        run_program_privatePay('12')

    def clickedDockingStudentsPrivatePay(self):
        run_docking_invoices_privatePay()

    def clickedPDFCreatorPrivatePay(self):
        excel_to_pdf_PrivatePay()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
    # runProgram('11')
    # run_program_elearning('11')
    # run_program_privatePay('11')
