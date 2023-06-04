#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 14:40:20 2022

@author: Alexander Mikhailov
"""
from pathlib import Path
from zipfile import ZipFile

# =============================================================================
# Separate Procedure
# =============================================================================
from win32com.client.gencache import EnsureDispatch

PATH_SRC = '/media/green-machine/KINGSTON'
PATH_EXP = '/media/green-machine/KINGSTON'


def compare_docx(file_control, file_test):
    Application = EnsureDispatch('Word.Application')
    Document = Application.Documents.Add()
    Application.CompareDocuments(Application.Documents.Open(Path(PATH_SRC).joinpath(file_control)),
                                 Application.Documents.Open(Path(PATH_SRC).joinpath(file_test)))
    # prevent that word opens itself
    Application.ActiveDocument.ActiveWindow.View.Type = 3
    Application.ActiveDocument.SaveAs(
        Path(PATH_EXP).joinpath('compared_docx.docx'))
    Application.Quit()


# =============================================================================
# Separate Procedure
# =============================================================================
with ZipFile('TextReview.zip') as archive_control:
    file_control = archive_control.read('Meeting Minutes 2014-11-15.docx')
with ZipFile('TextReview.zip') as archive_test:
    file_test = archive_test.read('Meeting Minutes 2014-11-15.docx')
Application = EnsureDispatch('Word.Application')
Document = Application.Documents.Add()
Application.CompareDocuments(Application.Documents.Open(Path(PATH_SRC).joinpath(file_control)),
                             Application.Documents.Open(Path(PATH_SRC).joinpath(file_test)))
# prevent that word opens itself
Application.ActiveDocument.ActiveWindow.View.Type = 3
Application.ActiveDocument.SaveAs(
    Path(PATH_EXP).joinpath('compared_docx.docx'))
Application.Quit()


compare_docx('Note INTH12 2015-10-29.docx', 'Note INTH12 2015-11-01.docx')


# =============================================================================
# Model R1-4F
# =============================================================================
