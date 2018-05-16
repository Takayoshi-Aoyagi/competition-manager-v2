# -*- coding: utf-8 -*-

import argparse
import json
import os
import sys

BASE_DIR = os.path.join(os.path.dirname(__file__), os.pardir)
LIB_DIR = os.path.join(BASE_DIR, 'lib')

sys.path.append(LIB_DIR)
from excel_infile_reader import ExcelInfileReader
from excel_entrant_sheet_writer import ExcelEntrantSheetWriter
from merger import Marger

class ExcelImport:

    def __init__(self, input_dir, out_file, merge=None):
        self.input_dir = input_dir
        self.out_file = out_file
        self.merge = merge

    def _get_files(self):
        files = []
        for (_, _, filenames) in os.walk(self.input_dir):
            filepaths = map(lambda x: os.path.join(self.input_dir, x), filenames)
            filepaths = filter(lambda x: x.endswith('xlsx') or x.endswith('xls'), filepaths)
            files.extend(filepaths)
        return files
    
    def get_entrants(self):
        entrants = []
        file_paths = self._get_files()
        for file_path in file_paths:
            print file_path
            reader = ExcelInfileReader(file_path)
            _entrants = reader.read()
            entrants.extend(_entrants)
        for i, e in enumerate(entrants):
            e['id'] = i + 1
        return entrants

    def execute(self):
        entrants = self.get_entrants()
        if self.merge is not None:
            entrants = Marger.merge(entrants, self.merge)
        writer = ExcelEntrantSheetWriter(entrants)
        writer.write(self.out_file)

def get_merge_configs(merge_strs):
    confs = []
    for s in merge_strs:
        arr = s.split(':')
        conf = {
            'event': arr[0],
            'new_name': arr[1],
            'names': map(lambda x: x.decode('utf-8'), arr[2].split(','))
            }
        confs.append(conf)
    return confs

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('input_dir')
    parser.add_argument('out_file')
    parser.add_argument('--merge', action='append')
    args = parser.parse_args()
    merge_configs = get_merge_configs(args.merge)
    
    ei = ExcelImport(args.input_dir, args.out_file, merge_configs)
    ei.execute()
