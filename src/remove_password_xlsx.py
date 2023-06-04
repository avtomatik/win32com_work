from win32com.client import Dispatch


def remove_password_xlsx(file_name: str, pw_str: str = 'WillisG2N%') -> None:
    # =========================================================================
    # TODO: Try Bytes
    # =========================================================================
    xcl = Dispatch('Excel.Application')
    wb = xcl.Workbooks.Open(file_name, False, False, None, pw_str)
    xcl.DisplayAlerts = False
    wb.SaveAs(file_name, None, '', '')
    xcl.Quit()
