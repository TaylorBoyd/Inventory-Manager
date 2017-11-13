from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton, QAction, QFileDialog, QLineEdit, QMessageBox, QLabel, QCheckBox
import sys
import xlrd
import xlwt
import csv
import datetime
from InventoryManager import *


full_oil_list = full_rumple_list()

class window(QMainWindow):


    def __init__(self, full_rumplestilskin_list):

        super(window, self).__init__()
        self.setGeometry(150, 150, 300, 365)
        self.setFixedSize(280, 340)
        self.setWindowTitle("Inventory Manager")

        self.full_list = full_rumplestilskin_list

        extractAction = QAction("Quit", self)
        extractAction.setShortcut("Ctrl+X")
        extractAction.setStatusTip("Closes the program")
        extractAction.triggered.connect(self.close_application)

        self.statusBar()

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu("&File")
        fileMenu.addAction(extractAction)

        self.home()

    def home(self):

        self.name_box = QLineEdit(self)
        self.name_box.move(95, 60)
        self.name_box.resize(125, 25)

        self.name_label = QLabel(self)
        self.name_label.setText("Oil Name")
        self.name_label.setAlignment(Qt.AlignVCenter)
        self.name_label.move(50, 57)
        self.name_label.setFixedWidth(50)

        self.prd_box = QLineEdit(self)
        self.prd_box.move(95, 90)
        self.prd_box.resize(125, 25)

        self.prd_label = QLabel(self)
        self.prd_label.setText("Production Code")
        self.prd_label.setAlignment(Qt.AlignVCenter)
        self.prd_label.move(13, 87)
        self.prd_label.setFixedWidth(80)

        self.newer_date = QLineEdit(self)
        self.newer_date.move(95, 120)
        self.newer_date.resize(125, 25)

        self.new_date_label = QLabel(self)
        self.new_date_label.setText("Purchased After")
        self.new_date_label.setAlignment(Qt.AlignVCenter)
        self.new_date_label.move(15, 117)
        self.new_date_label.setFixedWidth(80)

        self.older_date = QLineEdit(self)
        self.older_date.move(95, 150)
        self.older_date.resize(125, 25)

        self.older_date_label = QLabel(self)
        self.older_date_label.setText("Purchased Before")
        self.older_date_label.setAlignment(Qt.AlignVCenter)
        self.older_date_label.move(6, 147)
        self.older_date_label.setFixedWidth(85)

        self.output_name = QLineEdit(self)
        self.output_name.move(95, 230)
        self.output_name.resize(125, 25)

        self.output_name_label = QLabel(self)
        self.output_name_label.setText("Output Name")
        self.output_name_label.setAlignment(Qt.AlignVCenter)
        self.output_name_label.move(16, 227)

        search_button = QPushButton("Search", self)
        search_button.resize(95, 25)
        search_button.move(125, 280)
        search_button = search_button.clicked.connect(self.search)

        self.b1 = QCheckBox("Show only in stock", self)
        self.b1.move(95, 175)
        self.b1.setFixedWidth(150)

        self.b2 = QCheckBox("Show only out of stock", self)
        self.b2.move(95, 200)
        self.b2.setFixedWidth(150)

        self.show()

    def search(self):

        filtered_list = []
        for i in self.full_list:
            filtered_list.append(i)

        if len(self.output_name.text()) == 0:
            self.output_error()
            return

        # ------------------------------------------
        # Filter for names and lot number
        # ------------------------------------------

        if len(self.name_box.text()) > 0:
            filtered_list = list(filter(lambda oil: self.name_box.text().lower() in oil[1].lower(), filtered_list))

        if len(self.prd_box.text()) > 0:
            filtered_list = list(filter(lambda oil: self.prd_box.text().lower()[-4:] in oil[0].lower(), filtered_list))

        # ------------------------------------------
        # Get Date Values and check for date filters
        # ------------------------------------------

        get_purchase_date(filtered_list)
        filtered_list = sort_by_date(filtered_list)

        if len(self.newer_date.text()) > 0:

            try:
                search_date_newer = datetime.date(int(self.newer_date.text()[6:]),
                                                  int(self.newer_date.text()[:2]),
                                                  int(self.newer_date.text()[3:5]))
            except (TypeError, ValueError):
                self.date_error()
                return

            filtered_list = list(filter(lambda oil: search_date_newer <= oil[7], filtered_list))

        if len(self.older_date.text()) > 0:

            try:
                search_date_older = datetime.date(int(self.older_date.text()[6:]),
                                                  int(self.older_date.text()[:2]),
                                                  int(self.older_date.text()[3:5]))
            except (TypeError, ValueError):
                self.date_error()
                return

            filtered_list = list(filter(lambda oil: search_date_older >= oil[7], filtered_list))

        # --------------------------------------
        # Get Stock Values and check for filters
        # --------------------------------------

        get_stock(filtered_list)

        if self.b1.isChecked():

            filtered_list = list(filter(lambda oil: oil[2] != "0 mL" or oil[3] != "0 mL", filtered_list))

        if self.b2.isChecked():

            filtered_list = list(filter(lambda oil: oil[2] == "0 mL" and oil[3] == "0 mL", filtered_list))

        # --------------------------------------
        # Finish and build file
        # --------------------------------------

        if len(filtered_list) == 0:
            self.error_window_no_matches()
            return

        if len(filtered_list) >= 100:
            x = self.too_many_oils()
            if x:
                pass
            else:
                return


        try:
            create_file(filtered_list, "{}.xls".format(self.output_name.text()))

            self.older_date.setText("")
            self.newer_date.setText("")
            self.name_box.setText("")
            self.prd_box.setText("")
            self.output_name.setText("")
            self.b1.setChecked(False)
            self.b2.setChecked(False)
            return
        except EnvironmentError:
            self.output_error()
            self.output_name.setText("")
            return

    def error_window_no_matches(self):

        choice = QMessageBox.question(self, 'Error',
                                       "No oils matched your search", QMessageBox.Ok)
        if choice == QMessageBox.Ok:
            pass

    def date_error(self):

        choice = QMessageBox.question(self, 'Error',
                                       "Date must be in MM/DD/YYYY format", QMessageBox.Ok)
        if choice == QMessageBox.Ok:
            pass

    def output_error(self):

        choice = QMessageBox.question(self, 'Error',
                                       "Invalid output name", QMessageBox.Ok)
        if choice == QMessageBox.Ok:
            pass

    def too_many_oils(self):

        choice = QMessageBox.question(self, 'Error',
                                       "Over 100 oils match your search, are you sure you want to continue?", QMessageBox.Yes, QMessageBox.No)
        if choice == QMessageBox.Yes:
            return True
        if choice == QMessageBox.No:
            return False


    def close_application(self):
        sys.exit()

def run():

    app = QApplication(sys.argv)
    GUI = window(full_oil_list)
    sys.exit(app.exec_())

run()