#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  6 18:00:37 2025

@author: alexandermikhailov
"""

import re
import zipfile

from core.config import ARCHIVE_NAME, BASE_PATH, PATH_DST
from core.funcs import compare_word_docs


def sanitize_filename(file_name: str) -> str:
    """Sanitize file name by replacing spaces and dots with hyphens."""
    return re.sub('[ .]', '-', file_name)


def is_word_document(file_name: str) -> bool:
    """Check if the file is a Microsoft Word document or other text-based formats."""
    word_extensions = (
        '.doc', '.docx',  # Microsoft Word documents
        '.rtf',           # Rich Text Format
        '.odt',           # OpenDocument Text
        '.dot', '.dotx',  # Word templates
        '.docm',          # Word macro-enabled documents
        '.fodt',          # OpenDocument Flat XML
    )
    return file_name.lower().endswith(word_extensions)


def extract_and_compare_files():
    """Extract files from the archive and compare them."""
    archive_path = BASE_PATH.joinpath(ARCHIVE_NAME)

    PATH_DST.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(archive_path) as archive:
        for file_name in archive.namelist():
            if not is_word_document(file_name):
                continue

            sanitized_name = sanitize_filename(file_name)
            try:
                print(f'Processing file: {file_name}')

                archive.extract(file_name, path=PATH_DST)

                extracted_file_path = PATH_DST.joinpath(file_name)
                original_file_path = BASE_PATH.joinpath(file_name)
                comparison_result_path = PATH_DST.joinpath(
                    f'compared_{sanitized_name}.docx'
                )

                compare_word_docs(
                    extracted_file_path,
                    original_file_path,
                    comparison_result_path
                )

                extracted_file_path.unlink()

            except Exception as e:
                print(f'Failed to process {file_name}: {e}')


if __name__ == '__main__':
    extract_and_compare_files()
