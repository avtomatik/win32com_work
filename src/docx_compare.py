#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 14:40:20 2022

@author: Alexander Mikhailov
"""
import zipfile

from core.config import (ARCHIVE_NAME, BASE_PATH, DATE, DATE_LEFT, DATE_RIGHT,
                         PATH_DST)
from core.funcs import docx_compare

file_control = f'Meeting Minutes {DATE}.docx'
file_test = f'Meeting Minutes {DATE}.docx'

PATH_CTRL = BASE_PATH.joinpath(file_control)
PATH_TEST = BASE_PATH.joinpath(file_test)
PATH_EXPR = PATH_DST.joinpath('compared_docx.docx')


docx_compare(PATH_CTRL, PATH_TEST, PATH_EXPR)

# =============================================================================
# Separate Procedure
# =============================================================================


FILE_NAME = f'Meeting Minutes {DATE}.docx'
with zipfile.ZipFile(ARCHIVE_NAME) as archive_control:
    file_control = archive_control.read(FILE_NAME)


with zipfile.ZipFile(ARCHIVE_NAME) as archive_test:
    file_test = archive_test.read(FILE_NAME)

# =============================================================================
# Separate Procedure
# =============================================================================
docx_compare(f'Note INTH12 {DATE_LEFT}.docx', f'Note INTH12 {DATE_RIGHT}.docx')
