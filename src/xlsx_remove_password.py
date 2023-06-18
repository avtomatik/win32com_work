from win32com.client.dynamic import Dispatch
from win32com.client.gencache import EnsureDispatch


def remove_password_xlsx(file_name: str, pw_str: str = 'WillisG2N%') -> None:
    # =========================================================================
    # TODO: Try Bytes
    # =========================================================================
    PROG_ID = 'Excel.Application'
    app = EnsureDispatch(PROG_ID) or Dispatch(PROG_ID)
    wb = app.Workbooks.Open(file_name, False, False, None, pw_str)
    app.DisplayAlerts = False
    wb.SaveAs(file_name, None, '', '')
    app.Quit()
