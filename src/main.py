# =============================================================================
# Procedure: Compare All Files from Archive with the Same on Flash Drive
# =============================================================================
import os
import re
import zipfile

from core.config import ARCHIVE_NAME, BASE_PATH, PATH_DST
from core.funcs import compare_word_docs


def main():

    _ = 0

    with zipfile.ZipFile(BASE_PATH.joinpath(ARCHIVE_NAME)) as archive:
        for file_name in archive.namelist():
            if not file_name.endswith('.txt'):
                _file_name = re.sub('[ .]', '-', file_name)
                try:
                    print(file_name)
                    archive.extract(file_name, path=PATH_DST)

                    PATH_CTRL = PATH_DST.joinpath(file_name)
                    PATH_TEST = BASE_PATH.joinpath(file_name)
                    PATH_EXPR = PATH_DST.joinpath(
                        f'compared{_:02n}-{_file_name}.docx')

                    compare_word_docs(PATH_CTRL, PATH_TEST, PATH_EXPR)
                    os.unlink(PATH_CTRL)
                except:
                    pass
