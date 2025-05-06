from pathlib import Path

from win32com.client.dynamic import Dispatch
from win32com.client.gencache import EnsureDispatch


def docx_compare(path_ctr: Path, path_tst: Path, path_dst: Path) -> None:
    # =========================================================================
    # TODO: Implement Context Manager
    # =========================================================================
    PROG_ID = 'Word.Application'
    app = EnsureDispatch(PROG_ID) or Dispatch(PROG_ID)
    Document = app.Documents.Add()
    app.CompareDocuments(
        app.Documents.Open(path_ctr),
        app.Documents.Open(path_tst)
    )
    # =========================================================================
    # prevent that word opens itself
    # =========================================================================
    app.ActiveDocument.ActiveWindow.View.Type = 3
    app.ActiveDocument.SaveAs(path_dst)
    app.Quit()
    del app, Document
