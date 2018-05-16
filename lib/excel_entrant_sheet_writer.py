# -*- coding: utf-8 -*-

import collections

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

    def _create_sheet_data(self, entrants):
        rows = []
        for entrant in entrants:
            row = self._get_all_entrant_row_data(entrant)
            rows.append(row)
        return rows

    def _get_uniq_data(self, column_name):
        keys = list(set(map(lambda x: x[column_name], self.entrants)))
        keys.sort()
        return keys
    
    def _create_event_sheet_data(self, data, prefix, key_name):
        keys = self._get_uniq_data(key_name)
        for key in keys:
            _entrants = filter(lambda x: x[key_name] == key, self.entrants)
            sheet_name = prefix + " " + key
            data[sheet_name] = self._create_sheet_data(_entrants)
        
    def _create_data(self):
        data = collections.OrderedDict()
        data[u'参加者一覧'] = self._create_sheet_data(self.entrants)
        self._create_event_sheet_data(data, 'M', 'massogi')
        self._create_event_sheet_data(data, 'T', 'tul')
        return data
        
    def write(self, file_path):
        data = self._create_data()
        writer = ExcelWriter(file_path)
        writer.write(data)
