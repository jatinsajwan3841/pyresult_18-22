from openpyxl import load_workbook
from openpyxl import Workbook

letter = 'a'


class result:
    def __init__(self):
        self.wb = Workbook()
        self.sheet = self.wb.active
        while True:
            try:
                n = str(input("Enter your full name \n"))
                self.name = n.lower()
                self.clg_id = int(input("enter your college ID \n"))
                break
            except ValueError:
                print("mzak baad me,sahi likho")

        while True:
            chc = input(
                "Choose your branch,\n Enter 1 for CS \t 2 for EE \t 3 for ME \t 4 for CE \t 5 for BT \t 6 for EC")
            if chc == '1':
                self.branch = 'C S'
                break
            elif chc == '2':
                self.branch = 'EE'
                break
            elif chc == '3':
                self.branch = 'ME'
                break
            elif chc == '4':
                self.branch = 'CE'
                break
            elif chc == '5':
                self.branch = 'BT'
                break
            elif chc == '6':
                self.branch = 'EC'
                break
            else:
                print("mze baad me lena sahi enter kro")

    def select(self, sem):

        self.sem = sem

        while True:

            if self.sem == 1:
                self.sem1 = load_workbook('dat\\B. TECH. I SEM DEC 18.xlsx', data_only=True)
                self.head = 1
                self.body = 4
                self.currentSheet1 = self.sem1[self.branch]
            elif self.sem == 2:
                self.sem2 = load_workbook('dat\\B. TECH. II SEM JUNE 2019.xlsx', data_only=True)
                self.head = 7
                self.body = 10
                self.currentSheet1 = self.sem2[self.branch]
            elif self.sem == 3:
                self.sem3 = load_workbook('dat\\B. TECH. III SEM DECEMBER 2019.xlsx', data_only=True)
                self.head = 13
                self.body = 16
                self.currentSheet1 = self.sem3[self.branch]

            for row in range(1, self.currentSheet1.max_row + 1):
                for column in "DE":
                    self.cell_name = "{}{}".format(column, row)
                    ex = self.currentSheet1[self.cell_name].value
                    if ex != None and type(ex) != int:
                        ex = ex.strip()
                        ex = ex.lower()
                    if ex == self.clg_id or ex == self.name:
                        global letter
                        letter = row
                        break

            if letter == 'a':
                print('Either you are LE, wrna bahut galat input maara tune')
                break

            for row in range(4, 7):
                for column in range(1, self.currentSheet1.max_column + 1):  # Here you can add or reduce the columns
                    self.sheet.cell(self.head, column).value = self.currentSheet1.cell(row, column).value
                self.head = self.head + 1

            for column in range(1, self.currentSheet1.max_column + 1):  # Here you can add or reduce the columns
                self.sheet.cell(self.body, column).value = self.currentSheet1.cell(letter, column).value

            self.wb.save("result.xlsx")
            break

    def Display(self):
        for row in range(1, self.sheet.max_row + 1):
            for column in range(1, self.sheet.max_column + 1):
                if self.sheet.cell(row, column).value is None:
                    print("\t", end='')
                if self.sheet.cell(row, column).value is not None:
                    print(self.sheet.cell(row, column).value, "\t", end='')
            print()

    def clear_screen(self):
        import os
        os.system("cls")
