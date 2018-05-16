# -*- coding: utf-8 -*-

from openpyxl import load_workbook

def uni(row, index):
    cell = row[index]
    if cell is None or cell.value is None:
        return None
    return unicode(cell.value)

class ExcelInfileReader:

    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = load_workbook(filename=self.file_path)

    def _read_by_row(self, row):
        rowlen = 8
        if rowlen > len(row):
            return None
        entrant = {
            "name": uni(row, 1),
            "grade": uni(row, 3),
            "dojo": uni(row, 4),
            "tul": uni(row, 5),
            "massogi": uni(row, 8),
            "special": uni(row, 9),
            "kana": uni(row, 13),
            'fname': self.file_path
        }
        
        if entrant['name'] is None or entrant['name'] == u'出場選手' or entrant['name'] == u'審判員':
            return None
        if entrant['dojo'] is None:
            return None
        return entrant
        
    def _read_by_sheet(self, sheet):
        entrants = []
        for row in sheet.iter_rows():
            entrant = self._read_by_row(row)
            if entrant is None:
                continue
            entrants.append(entrant)
        return entrants
            
    def read(self):
        entrants = []
        for sheet_name in self.wb.sheetnames:
            sheet = self.wb[sheet_name]
            _entrants = self._read_by_sheet(sheet)
            entrants.extend(_entrants)
        return entrants
            
