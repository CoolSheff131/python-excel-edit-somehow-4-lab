import sys

import openpyxl
import xlrd
from PyQt5.QtCore import Qt

from design import Ui_Dialog

from PyQt5.QtWidgets import (
    QApplication, QDialog, QMainWindow, QMessageBox, QTableWidgetItem
)
from PyQt5.uic import loadUi
import pandas as pd  # pip install pandas
import numpy as np
from xlrd import open_workbook
from openpyxl import load_workbook, Workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Font, Border

class Window(QMainWindow, Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.wb2 = None
        self.columns = None
        self.setupUi(self)

        self.loadRazdachaBtn.clicked.connect(self.loadRazdacha)
        self.loadHappyBtn.clicked.connect(self.loadHappyData)
        self.saveHappyBtn.clicked.connect(self.saveHappyData)
        self.saveRazdachaBtn.clicked.connect(self.saveRazdacha)

        self.table.cellClicked.connect(self.cellClicked)

        self.path1 = './self.path1.xlsx'

        self.path2 = './2.xlsx'


    def cellClicked(self, row, column):
        print(row, column)
        while (self.table_3.rowCount() > 0):
            self.table_3.removeRow(0)

        self.table_3.setRowCount(self.table_2.rowCount())
        self.table_3.setColumnCount(self.table_2.columnCount())

        clickedRowText = self.table.item(row, 1).text()
        print(clickedRowText)
        self.label_3.setText(clickedRowText)
        lastRowInsertedIndex = 0
        for rowIndex in range(self.table_2.rowCount()):
            rowText = self.table_2.item(rowIndex, 0).text()
            print(rowText)
            print(rowText == clickedRowText)
            if rowText == clickedRowText:
                for columnIndex in range(self.table_2.columnCount()):
                    row = self.table_2.item(rowIndex, columnIndex).text()
                    tableItem = QTableWidgetItem(row)
                    self.table_3.setItem(lastRowInsertedIndex, columnIndex, tableItem)
                lastRowInsertedIndex += 1

        self.table_3.setColumnWidth(2, 300)

    def saveRazdacha(self):

        # df = pd.DataFrame()
        # for row in range(self.table.rowCount()):
        #     for col in range(self.table.columnCount()):
        #         print(self.table.item(row, col).text())
        #         df.at[row - 2, col] = self.table.item(row, col).text()
        # df.to_excel('./Dummy File XYZ.xlsx', index=False,header=False)

        filename = '2self.path1.xlsx'
        table = self.table
        print('Save razdacha')
        wb = Workbook()
        sheet = wb.active
        for row in range(table.rowCount()):
            for column in range(table.columnCount()):
                text = self.table.item(row, column).text()

                try:
                    print(row, column)
                    text = str(table.item(row, column).text())
                    print(text)
                    sheet.cell(row+1, column+1).value = text

                    background = table.item(row, column).background()
                    redFill = PatternFill(start_color='FFFFFFFF',
                                          end_color='FFFFFFFF',
                                          fill_type='solid')

                    if (background == Qt.yellow):
                        redFill = PatternFill(start_color='FFFFFF00',
                                              end_color='FFFFFF00',
                                              fill_type='solid')
                    elif (background == Qt.red):
                        redFill = PatternFill(start_color='FFFF0000',
                                              end_color='FFFF0000',
                                              fill_type='solid')
                    elif (background == Qt.green):
                        redFill = PatternFill(start_color='FF00ff00',
                                              end_color='FF00ff00',
                                              fill_type='solid')

                    sheet.cell(row+1, column+1).fill = redFill

                    # colorHex = cell_obj.fill.start_color.index
                    # if (colorHex == 'FFFFFF00'):
                    #     tableItem.setBackground(Qt.yellow);
                    # elif (colorHex == 7):
                    #     tableItem.setBackground(Qt.red);
                    # elif (colorHex == 'FF92D050'):
                    #     tableItem.setBackground(Qt.green);
                except AttributeError:
                    pass

        wb.save(filename)

    def loadRazdacha(self):
        while (self.table_3.rowCount() > 0):
            self.table_3.removeRow(0)
        sheetName = 'дост26.12 (2)'

        wb_obj = openpyxl.load_workbook(self.path1)
        self.wb1 = wb_obj
        sheet_obj = wb_obj[sheetName]
        m_row = sheet_obj.max_row
        m_column = sheet_obj.max_row
        self.table.setRowCount(m_row)
        self.table.setColumnCount(m_column)
        for i in range(1, m_row + 1):
            for j in range(1, m_column +1 ):
                cell_obj = sheet_obj.cell(row=i, column=j)
                tableItem = QTableWidgetItem(str(cell_obj.value or ''))
                self.table.setItem(i -1, j -1, tableItem)
                colorHex = cell_obj.fill.start_color.index
                if (colorHex == 'FFFFFF00'):
                    tableItem.setBackground(Qt.yellow);
                elif (colorHex == 7): 
                    tableItem.setBackground(Qt.red);
                elif (colorHex == 'FF92D050'):
                    tableItem.setBackground(Qt.green);
        self.table.setColumnWidth(2, 300)

    def saveHappyData(self):
        # self.wb2.save(self.path2)
        print('Save happy')

        self.export(self, self.path2, self.table_2)

    def loadHappyData(self):
        sheetName = '135-27.12'
        wb_obj = openpyxl.load_workbook(self.path2)
        self.wb2 = wb_obj
        sheet_obj = wb_obj[sheetName]
        m_row = sheet_obj.max_row
        m_column = sheet_obj.max_row
        self.table_2.setRowCount(m_row)
        self.table_2.setColumnCount(m_column)
        for i in range(1, m_row + 1):
            for j in range(1, m_column + 1):
                cell_obj = sheet_obj.cell(row=i, column=j)
                tableItem = QTableWidgetItem(str(cell_obj.value or ''))
                self.table_2.setItem(i - 1, j - 1, tableItem)
                colorHex = cell_obj.fill.start_color.index
                if (colorHex == 'FFFFFF00'):
                    tableItem.setBackground(Qt.yellow);
                elif (colorHex == 7):
                    tableItem.setBackground(Qt.red);
                elif (colorHex == 'FF92D050'):
                    tableItem.setBackground(Qt.green);
        self.table_2.setColumnWidth(2, 300)

    def export(self, filename, table):
        print('saave')
        # filename, filter = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', '','Excel files (*.xlsx)')



if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
