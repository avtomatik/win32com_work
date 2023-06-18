# =============================================================================
# Procedure: Compare All Files from Archive with the Same on Flash Drive
# =============================================================================
import os
import re
from pathlib import Path
from zipfile import ZipFile

from core.funcs import docx_compare


def main():
    PATH_SRC = '/media/green-machine/KINGSTON'
    PATH_EXP = '/media/green-machine/KINGSTON'

    _ = 0

    ARCHIVE_NAME = 'TextReview.zip'

    with ZipFile(Path(PATH_SRC).joinpath(ARCHIVE_NAME)) as archive:
        for file_name in archive.namelist():
            if not file_name.endswith('.txt'):
                _file_name = re.sub('[ .]', '-', file_name)
                try:
                    print(file_name)
                    archive.extract(file_name, path=PATH_EXP)

                    PATH_CTRL = Path(PATH_EXP).joinpath(file_name)
                    PATH_TEST = Path(PATH_SRC).joinpath(file_name)
                    PATH_EXPR = Path(PATH_EXP).joinpath(
                        f'compared{_:02n}-{_file_name}.docx')

                    docx_compare(PATH_CTRL, PATH_TEST, PATH_EXPR)
                    os.unlink(PATH_CTRL)
                except:
                    pass
