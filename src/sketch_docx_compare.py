from pathlib import Path
from zipfile import ZipFile

from win32com.client.gencache import EnsureDispatch


def docx_compare(file_control: str, file_test: str) -> None:
    PATH = 'G:\\G_Energy\\A MIKHAILOV\\miscellaneous'
    Application = EnsureDispatch('Word.Application')
    Document = Application.Documents.Add()
    Application.CompareDocuments(
        Application.Documents.Open(
            Path(PATH).joinpath(f'{file_control}.docx')),
        Application.Documents.Open(Path(PATH).joinpath(f'{file_test}.docx'))
    )
    # =============================================================================
    # prevent that word opens itself
    # =============================================================================
    Application.ActiveDocument.ActiveWindow.View.Type = 3
    Application.ActiveDocument.SaveAs(Path(PATH).joinpath('Compared.docx'))
    Application.Quit()


# =============================================================================
# Separate Procedure
# =============================================================================

archive_control = ZipFile('Processing Text Review.zip')
file_control = archive_control.read('Meeting Minutes 2014-11-15.docx')
archive_test = ZipFile('Processing Text Review.zip')
file_test = archive_test.read('Meeting Minutes 2014-11-15.docx')
Application = EnsureDispatch('Word.Application')
Document = Application.Documents.Add()
Application.CompareDocuments(
    Application.Documents.Open(file_control),
    Application.Documents.Open(file_test)
)
# =============================================================================
# prevent that word opens itself
# =============================================================================
Application.ActiveDocument.ActiveWindow.View.Type = 3
Application.ActiveDocument.SaveAs(Path(PATH).joinpath('Compared.docx'))
Application.Quit()
docx_compare('CLASSIFIED Bluestream PD', 'CLASSIFIED Bluestream Terrorism')
