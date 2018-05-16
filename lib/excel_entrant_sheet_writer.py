# -*- coding: utf-8 -*-

from excel_writer import ExcelWriter

class ExcelEntrantSheetWriter:

    def __init__(self, entrants):
        self.entrants = entrants

    def _get_all_entrant_row_data(self, entrant):
        keys = ("name", "grade", "dojo", "tul", "massogi", "special", "kana", "fname")
        row = []
        for key in keys:
            row.append(entrant[key])
        return row

    def _create_all_entrants_sheet_data(self):
        rows = []
        for entrant in self.entrants:
            row = self._get_all_entrant_row_data(entrant)
            rows.append(row)
        return rows

    def _create_data(self):
        data = {}
        data[u'参加者一覧'] = self._create_all_entrants_sheet_data()        
        return data
        
    def write(self, file_path):
        data = self._create_data()
        writer = ExcelWriter(file_path)
        writer.write(data)
