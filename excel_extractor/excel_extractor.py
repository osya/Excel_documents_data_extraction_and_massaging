#!/usr/bin/env python
# -*- coding: utf-8 -*-
from abc import abstractmethod
import pandas as pd
from collections import OrderedDict


class ExcelExtractor(object):
    def extract(self):
        col_map = OrderedDict(zip(self._col_names, self.col_names_cur))
        res = self._df[[col for col in col_map.values() if col]]
        res.columns = [key for key in col_map if col_map[key]]
        return res

    def __init__(self, file_name, sheet_name, skip_rows=None):
        self._col_names = [u'Lot number out of total', u'Quantity', u'Bottle size', u'Vintage', u'Name', u'Designation',
                           u'Producer', u'Levels', u'Low', u'High', u'Sale Price', u'Region', u'Score'
                           ]
        self._df = pd.read_excel(open(file_name, 'rb'), sheetname=sheet_name, skiprows=skip_rows)

    @staticmethod
    def factory(type, file_name):
        if 1 == type:
            return ExcelExtractor1(file_name)
        if 2 == type:
            return ExcelExtractor2(file_name)
        if 3 == type:
            return ExcelExtractor3(file_name)
        if 4 == type:
            return ExcelExtractor4(file_name)
        if 5 == type:
            return ExcelExtractor5(file_name)
        assert 0, 'Bad ExcelExtractor type: %s' % type


class ExcelExtractor1(ExcelExtractor):
    def __init__(self, file_name):
        super(ExcelExtractor1, self).__init__(file_name, 'qryCatalogExcel')
        self.col_names_cur = [u'LotNo', u'Quantity', None, u'Vintage', None, u'Designation', u'Producer', u'Levels',
                              u'Low', u'High', None, u'RegionDescription', u'ItemWineScore']


class ExcelExtractor2(ExcelExtractor):
    def __init__(self, file_name):
        super(ExcelExtractor2, self).__init__(file_name, 'Sheet1', skip_rows=1)
        self.col_names_cur = [u'Lot No.', u'Quantity', None, None, None, None, None, None, u'Estimate Low $USD',
                              u'Estimate High $USD', None, None, None]


class ExcelExtractor3(ExcelExtractor):
    def __init__(self, file_name):
        super(ExcelExtractor3, self).__init__(file_name, 'Wine Online NYC_1-15 September', skip_rows=5)
        self.col_names_cur = [u'Lot', u'Qty', u'Size', u'Vintage', u'Title', None, None, None, u'Low Estimate',
                              u'High Estimate', None, None, None]


class ExcelExtractor4(ExcelExtractor):
    def __init__(self, file_name):
        super(ExcelExtractor4, self).__init__(file_name, 'Sheet1')
        self.col_names_cur = [u'Lot number', None, None, None, None, None, None, None, u'Low estimate',
                              u'High estimate', None, None, None]


class ExcelExtractor5(ExcelExtractor):
    def __init__(self, file_name):
        super(ExcelExtractor5, self).__init__(file_name, '0915NY')
        self.col_names_cur = [u'Lot', u'Qty', u'Size', u'Vintage', u'Description', None, None, None, u'Low', u'High',
                              None, u'Region', u'Score']
