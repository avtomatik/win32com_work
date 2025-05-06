#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  6 16:45:27 2025

@author: alexandermikhailov
"""

from pathlib import Path

from win32com.client import Dispatch
from win32com.client.gencache import EnsureDispatch

from core.config import PASSWORD


def remove_excel_password(file_path: Path, password: str = PASSWORD) -> None:
    """Removes the password protection from an Excel file and saves it.

    Args:
        file_path (Path): Path to the password-protected Excel file.
        password (str): Password for the Excel file. Default is PASSWORD.
    """
    EXCEL_PROG_ID = 'Excel.Application'

    app = EnsureDispatch(EXCEL_PROG_ID) or Dispatch(EXCEL_PROG_ID)
    app.DisplayAlerts = False

    try:
        workbook = app.Workbooks.Open(
            Filename=file_path,
            UpdateLinks=False,
            ReadOnly=False,
            Password=password
        )
        workbook.SaveAs(Filename=file_path, Password='', WriteResPassword='')
        workbook.Close(SaveChanges=True)
    finally:
        app.Quit()
