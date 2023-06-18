from win32com.client.dynamic import Dispatch
from win32com.client.gencache import EnsureDispatch


def docx_compare(PATH_CTRL, PATH_TEST, PATH_EXPR):
    # =========================================================================
    # TODO: Implement Context Manager
    # =========================================================================
    PROG_ID = 'Word.Application'
    app = EnsureDispatch(PROG_ID) or Dispatch(PROG_ID)
    Document = app.Documents.Add()
    app.CompareDocuments(
        app.Documents.Open(PATH_CTRL),
        app.Documents.Open(PATH_TEST)
    )
    # =========================================================================
    # prevent that word opens itself
    # =========================================================================
    app.ActiveDocument.ActiveWindow.View.Type = 3
    app.ActiveDocument.SaveAs(PATH_EXPR)
    app.Quit()
    del app, Document
