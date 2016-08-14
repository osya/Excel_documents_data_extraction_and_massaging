#!/usr/bin/env python
# -*- coding: utf-8 -*-
from excel_extractor import ExcelExtractor
import os

# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_1.xls'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_2.xlsx'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_3.xls'))
# file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_4.xlsx'))
file_path = os.path.abspath(os.path.join(os.getcwd(), '..', 'doc_samples', 'template_5.xlsx'))

ee = ExcelExtractor.factory(5, file_path)
ee.extract()
