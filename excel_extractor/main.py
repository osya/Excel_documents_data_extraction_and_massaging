#!/usr/bin/env python
# -*- coding: utf-8 -*-
from excel_extractor import ExcelExtractor
import os

file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_1.xls'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_2.xlsx'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_3.xls'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_4.xlsx'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_5.xlsx'))


# dfs = {sheet_name: xl_file.parse(sheet_name)
#           for sheet_name in xl_file.sheet_names}

ee = ExcelExtractor.factory(1, file_path)
ee.extract()
