import xlrd
import xlwt
import csv
import datetime
from operator import itemgetter


def full_rumple_list():

    cell = 0
    oil_list = []

    try:
        book = xlrd.open_workbook("Rumplestilskin.xls")
        sh = book.sheet_by_index(0)
    except NameError:
        print("You are missing the file Rumplestilskin.xls")
        return

    while sh.cell_value(rowx=cell, colx=0) != "END":

        if len(sh.cell_value(rowx=cell, colx=1)) > 0:

            temp_oil = []
            temp_oil.append(sh.cell_value(rowx=cell, colx=1))  # 0 Lot#
            temp_oil.append(sh.cell_value(rowx=cell, colx=0))  # 1 Oil Name
            temp_oil.append("")                                # 2 Main Stock
            temp_oil.append("")                                # 3 Back Stock
            temp_oil.append("")                                # 4 Purchase date (str) for display
            temp_oil.append(sh.cell_value(rowx=cell, colx=5))  # 5 Country of origin
            temp_oil.append(sh.cell_value(rowx=cell, colx=8))  # 6 Cultivation type
            temp_oil.append("")                                # 7 Date time object for sorting

            oil_list.append(temp_oil)
        cell += 1

    return oil_list

def create_oil_list(oil_name, product_code):

    oil_list = []
    count = 0
    cell = 0

    try:
        book = xlrd.open_workbook("Rumplestilskin.xls")
        sh = book.sheet_by_index(0)
    except NameError:
        print("You are missing the file Rumplestilskin.xls")
        return

    while sh.cell_value(rowx=cell, colx=0) != "END":

        if oil_name.lower() in sh.cell_value(rowx=cell, colx=0).lower():
            if product_code.lower() in sh.cell_value(rowx=cell, colx=1)[-4:].lower():
                temp_oil = []
                temp_oil.append(sh.cell_value(rowx=cell, colx=1))    # 0 Lot#
                temp_oil.append(sh.cell_value(rowx=cell, colx=0))    # 1 Oil Name
                temp_oil.append("")                                  # 2 Main Stock
                temp_oil.append("")                                  # 3 Back Stock
                temp_oil.append("")                                  # 4 Purchase date (str) for display
                temp_oil.append(sh.cell_value(rowx=cell, colx=5))    # 5 Country of origin
                temp_oil.append(sh.cell_value(rowx=cell, colx=8))    # 6 Cultivation type
                temp_oil.append("")                                  # 7 Date time object for sorting

                oil_list.append(temp_oil)
                count += 1

        cell += 1

    if count == 0:
        return [[]]
    else:
        return oil_list

def get_stock(oil_list):

    with open("current stock.csv") as cs:
        current_stock = csv.reader(cs, delimiter=",")

        for row in current_stock:
            for oil in oil_list:
                if oil[0].lower() in row[3].lower():
                    oil[2] = "{} mL".format(row[2])

    with open("backstock 1.csv") as bs1:
        backstock1 = csv.reader(bs1, delimiter=",")

        for row in backstock1:
            for oil in oil_list:
                if oil[0].lower() in row[3].lower():
                    oil[3] = "{} mL".format(row[2])

    with open("backstock 2.csv") as bs2:
        backstock2 = csv.reader(bs2, delimiter=",")

        for row in backstock2:
            for oil in oil_list:
                if oil[0].lower() in row[3].lower():
                    oil[3] = "{} mL".format(row[2])

    with open("backstock 3.csv") as bs3:
        backstock3 = csv.reader(bs3, delimiter=",")

        for row in backstock3:
            for oil in oil_list:
                if oil[0].lower() in row[3].lower():
                    oil[2] = "{} mL".format(row[2])

    for oil in oil_list:
        if len(oil[2]) == 0:
            oil[2] = "0 mL"
        if len(oil[3]) == 0:
            oil[3] = "0 mL"

    return(oil_list)

def get_purchase_date(oil_list):

    month_list = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6,
                  "G": 7, "H": 8, "I": 9, "J": 10, "K": 11, "L": 12}
    year_list = {"1": 2021, "2": 2022, "3": 2013, "4": 2014, "5": 2015, "6": 2016,
                 "7": 2017, "8": 2018, "9": 2019, "0": 2020}

    for oil in oil_list:
        try:
            month = month_list[oil[0][2:3].upper()]
            year = year_list[oil[0][5:6]]
            day = int(oil[0][3:5])
            oil[4] = ("{}/{}/{}".format(month, day, year))
            oil[7] = datetime.date(year, month, day)

        except KeyError:
            oil[4] = "Invalid Date"
            oil[7] = datetime.date(2001, 1, 1)
        except ValueError:
            oil[4] = "Invalid Date"
            oil[7] = datetime.date(2001, 1, 1)

    return oil_list

def sort_by_date(oil_list):

    temp_list = oil_list
    oil_list = sorted(temp_list, key=itemgetter(7), reverse=True)

    return oil_list

def create_file(oil_list, output_name):

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Inventory Info")

    Titles = ["Lot #", "Oil Name", "Main Stock", "Back Stock", "Purchase Date", "Country of Origin", "Cultivation Type"]
    column = 0
    for i in Titles:
        ws.write(0, column, i, xlwt.easyxf("align: horiz center; font: bold on; borders: bottom thin"))
        column += 1

    columns = (len(oil_list[0]) - 1)
    rows = len(oil_list)
    for i in range(columns):
        for j in range(rows):
            x = (oil_list[j][i])
            ws.write(j+2, i, x, xlwt.easyxf("align: horiz right"))

    for i in range(columns):
        ws.col(i).width = 256*25

    wb.save(output_name)

    return


if __name__ == "__main__":

    oil_name = input("Please type and oil to create an inventory sheet: ")
    out_put_name = oil_name + ".xls"

    while True:
        
        oil_list = create_oil_list(oil_name)

        if oil_list != False:
            get_stock(oil_list)
            get_purchase_date(oil_list)
            oil_list = sort_by_date(oil_list)
            create_file(oil_list, out_put_name)
        
            print("Complete!")
            print("")
            choice = input("Type 'q' to quit or type a new oil to create a sheet for: ")
        else:
            print("That oil does not exist!")
            print("")
            choice = input("Type 'q' to quit or type a new oil to create a sheet for: ")

        if choice.lower() == "q":
            break
        else:
            oil_name = choice.lower()
            out_put_name = oil_name + ".xls"
