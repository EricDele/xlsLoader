#! /usr/bin/env python
# -*- coding: utf-8 -*-

import json
from openpyxl import load_workbook

class XlsLoader:

    def __init__(self, xls_file, config_file):
        self.xls_file = xls_file
        with open(config_file) as jsonFile:
            self.config_file = json.load(jsonFile)

    def load_file(self):
        wb = load_workbook(self.xls_file, data_only=True)
        ws = wb.get_sheet_by_name(self.config_file['sheet-name'])
        row = int(self.config_file['cellule-origine']['row'])
        col = int(self.config_file['cellule-origine']['col'])
        rowMax = int(self.config_file['row-max'])
