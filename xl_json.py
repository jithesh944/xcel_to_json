#! /Volumes/Media/venv/myDjango_venv/bin/python3

from typing import Dict
from openpyxl import load_workbook
import os,json
from itertools import islice
from collections import OrderedDict

# loading the excel file from the same location.
xcel_file = load_workbook('Untitled.xlsx')

# Extracting the sheets names
xcel_sheet_names = xcel_file.sheetnames

# Multiple sheet reading logic
for (i, sheet) in enumerate(xcel_sheet_names):
    header_row = []
    ws = xcel_file[sheet]
    m_row = xcel_file[sheet].max_row
    m_col = xcel_file[sheet].max_column

    # print(m_row, m_col)



   