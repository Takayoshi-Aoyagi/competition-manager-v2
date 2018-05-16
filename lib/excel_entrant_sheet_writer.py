# -*- coding: utf-8 -*-

import collections

from excel_writer import ExcelWriter

class EntrantSheet:

    @staticmethod
    def get_column_defs():
        keys = (
            ('id', 'ゼッケン'),
            ("name", '氏名'),
            ("kana", 'ふりがな'),
            ("gender", '性別'),
            ("grade", '級・段'),
            ("dojo", '道場'),
            ("tul", 'トゥル'),
            ("massogi", 'マッソギ'),
            ("special", 'スペシャル'),
            ("fname", 'ファイル名')
        )
        return keys

class ExcelEntrantSheetWriter:

    def __init__(self, entrants):
        self.entrants = entrants
        self.column_defs = EntrantSheet.get_column_defs()
        self.header = map(lambda x: x[1], self.column_defs)

    def _get_all_entrant_row_data(self, entrant):
        keys = map(lambda x: x[0], self.column_defs)
        row = []
        for key in keys:
            row.append(entrant[key])
        return row

    def _create_sheet_data(self, entrants):
        rows = []
        rows.append(self.header)
        for entrant in entrants:
            row = self._get_all_entrant_row_data(entrant)
            rows.append(row)
        return rows

    def _get_uniq_data(self, column_name):
        keys = list(set(map(lambda x: x[column_name], self.entrants)))
        keys = filter(lambda x: x is not None, keys)
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
