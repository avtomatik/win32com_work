#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 14:40:20 2022

@author: Alexander Mikhailov
"""
import datetime
from pathlib import Path
from zipfile import ZipFile

from core.funcs import docx_compare

DATE = datetime.date(2014, 11, 15)

PATH_SRC = '/media/green-machine/KINGSTON'
PATH_EXP = '/media/green-machine/KINGSTON'
file_control = f'Meeting Minutes {DATE}.docx'
file_test = f'Meeting Minutes {DATE}.docx'
PATH_CTRL = Path(PATH_SRC).joinpath(file_control)
PATH_TEST = Path(PATH_SRC).joinpath(file_test)
PATH_EXPR = Path(PATH_EXP).joinpath('compared_docx.docx')


docx_compare(PATH_CTRL, PATH_TEST, PATH_EXPR)

# =============================================================================
# Separate Procedure
# =============================================================================

ARCHIVE_NAME = 'TextReview.zip' or 'Processing Text Review.zip'
DATE = datetime.date(2014, 11, 15)
FILE_NAME = f'Meeting Minutes {DATE}.docx'
with ZipFile(ARCHIVE_NAME) as archive_control:
    file_control = archive_control.read(FILE_NAME)


with ZipFile(ARCHIVE_NAME) as archive_test:
    file_test = archive_test.read(FILE_NAME)

# =============================================================================
# Separate Procedure
# =============================================================================
DATE_LEFT = datetime.date(2015, 10, 29)
DATE_RIGHT = datetime.date(2015, 11, 1)
docx_compare(f'Note INTH12 {DATE_LEFT}.docx', f'Note INTH12 {DATE_RIGHT}.docx')
