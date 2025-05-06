#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  6 17:30:47 2025

@author: alexandermikhailov
"""

from contextlib import contextmanager
from pathlib import Path

from win32com.client.gencache import EnsureDispatch


@contextmanager
def word_app():
    """Context manager to safely start and close Word application."""
    app = EnsureDispatch('Word.Application')
    app.Visible = False
    try:
        yield app
    finally:
        app.Quit()
        del app


def compare_word_docs(path_ctr: Path, path_tst: Path, path_dst: Path) -> None:
    """
    Compares two Word documents and saves the result.

    :param path_ctr: Path to the control/original document.
    :param path_tst: Path to the test/modified document.
    :param path_dst: Path to save the comparison result.
    """
    with word_app() as app:
        doc_ctr = app.Documents.Open(str(path_ctr))
        doc_tst = app.Documents.Open(str(path_tst))

        doc_comparison = app.CompareDocuments(doc_ctr, doc_tst)
        doc_comparison.ActiveWindow.View.Type = 3
        doc_comparison.SaveAs(str(path_dst))

        doc_comparison.Close(False)
        doc_tst.Close(False)
        doc_ctr.Close(False)
