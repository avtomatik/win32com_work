# =============================================================================
# Procedure: Compare All Files from Archive with the Same on Flash Drive
# =============================================================================
import os
import re
from pathlib import Path
from zipfile import ZipFile

from win32com.client.dynamic import Dispatch
from win32com.client.gencache import EnsureDispatch

from archives.src.archives_manipulations import _


def main():
    PATH = 'D:'
    PATH_EXP = '/media/green-machine/KINGSTON'

    with ZipFile(Path(PATH).joinpath('TextReview.zip')) as archive:
        for file_name in archive.namelist():
            if not file_name.endswith('.txt'):
                _file_name = re.sub('[ .]', '-', file_name)
                try:
                    print(file_name)
                    archive.extract(file_name, path=PATH_EXP)
                    Application = EnsureDispatch('Word.Application')
                    Application = Dispatch('Word.Application')
                    Document = Application.Documents.Add()
                    Application.CompareDocuments(
                        Application.Documents.Open(
                            Path(PATH_EXP).joinpath(file_name)),
                        Application.Documents.Open(
                            Path(PATH).joinpath(file_name))
                    )
                # =================================================================
                # prevent that word opens itself
                # =================================================================
                    Application.ActiveDocument.ActiveWindow.View.Type = 3
                    Application.ActiveDocument.SaveAs(
                        os.path.join(
                            PATH_EXP, 'compared{:02n}-{}.docx'.format(_, _file_name))
                    )
                    Application.Quit()
                    del Application, Document
                    os.unlink(Path(PATH_EXP).joinpath(file_name))
                except:
                    pass
