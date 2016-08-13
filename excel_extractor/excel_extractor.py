#!/usr/bin/env python
# -*- coding: utf-8 -*-
from abc import abstractmethod
import pandas as pd
from collections import OrderedDict


class ExcelExtractor(object):
    @abstractmethod
    def extract(self):
        return pd.ExcelFile(self._file_name)

    def __init__(self, file_name):
        self._file_name = file_name
        self._col_names = ['Lot number out of total', 'Quantity', 'Bottle size', 'Vintage', 'Name', 'Designation',
                           'Producer', 'Levels', 'Low', 'High', 'Sale Price', 'Region', 'Score'
                           ]

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
    def extract(self):
        xl = super(ExcelExtractor1, self).extract().parse('qryCatalogExcel')
        col_names1 = ['LotNo', 'Quantity', None, 'Vintage', None, 'Designation', 'Producer', 'Levels', 'Low', 'High',
                      None, 'RegionDescription', 'ItemWineScore']
        col_map = OrderedDict(zip(self._col_names, col_names1))
        res = xl[[col for col in col_map.values() if col]]
        res.columns = [key for key in col_map if col_map[key]]
        return res


class ExcelExtractor2(ExcelExtractor):
    def extract(self):
        xl = super(ExcelExtractor1, self).extract().parse('Sheet1')
        return xl


class ExcelExtractor3(ExcelExtractor):
    def extract(self):
        xl = super(ExcelExtractor1, self).extract().parse('Wine Online NYC_1-15 September')
        return xl


class ExcelExtractor4(ExcelExtractor):
    def extract(self):
        xl = super(ExcelExtractor1, self).extract().parse('Sheet1')
        return xl


class ExcelExtractor5(ExcelExtractor):
    def extract(self):
        xl = super(ExcelExtractor1, self).extract().parse('0915NY')
        return xl
