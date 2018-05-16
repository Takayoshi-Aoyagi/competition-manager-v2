# -*- coding: utf-8 -*-

import openpyxl

class ExcelWriter:

    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = openpyxl.Workbook()
        s0 = self.wb['Sheet']
        self.wb.remove(s0)

    def _create_row(self, sheet, row_index, row):
        col_indexes = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        for i, col in enumerate(row):
            col_index = col_indexes[i]
            addr = "{0}{1}".format(col_index, row_index)
            sheet[addr].value = col

    def _create_sheet(self, sheet_name, rows):
        sheet = self.wb.create_sheet(sheet_name)
        for i, row in enumerate(rows):
            row_index = i + 1
            self._create_row(sheet, row_index, row)
        
    def write(self, data):
        for (sheet_name, sheet_data) in data.items():
            self._create_sheet(sheet_name, sheet_data)
        self.wb.save(self.file_path)
