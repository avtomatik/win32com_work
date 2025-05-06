#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 14:40:20 2022

@author: Alexander Mikhailov
"""

import zipfile

from core.config import ARCHIVE_NAME, BASE_PATH, DATE

# =============================================================================
# Separate Procedure
# =============================================================================
file_name = f'Meeting Minutes {DATE}.docx'

with zipfile.ZipFile(BASE_PATH.joinpath(ARCHIVE_NAME)) as archive_ctr:
    doc_ctr = archive_ctr.read(file_name)


with zipfile.ZipFile(BASE_PATH.joinpath(ARCHIVE_NAME)) as archive_tst:
    doc_tst = archive_tst.read(file_name)
