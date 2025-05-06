#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 14:40:20 2022

@author: Alexander Mikhailov
"""

import zipfile

from core.config import (ARCHIVE_NAME, BASE_PATH, DATE, DATE_CTR, DATE_TST,
                         PATH_DST)
from core.funcs import compare_word_docs

name_ctr = f'Meeting Minutes {DATE}.docx'
name_tst = f'Meeting Minutes {DATE}.docx'
name_dst = 'compared_docx.docx'

path_ctr = BASE_PATH.joinpath(name_ctr)
path_tst = BASE_PATH.joinpath(name_tst)
path_dst = PATH_DST.joinpath(name_dst)

compare_word_docs(path_ctr, path_tst, path_dst)

# =============================================================================
# Separate Procedure
# =============================================================================
file_name = f'Meeting Minutes {DATE}.docx'

with zipfile.ZipFile(ARCHIVE_NAME) as archive_ctr:
    doc_ctr = archive_ctr.read(file_name)


with zipfile.ZipFile(ARCHIVE_NAME) as archive_tst:
    doc_tst = archive_tst.read(file_name)

# =============================================================================
# Separate Procedure
# =============================================================================
compare_word_docs(f'Note INTH12 {DATE_CTR}.docx',
                  f'Note INTH12 {DATE_TST}.docx')
