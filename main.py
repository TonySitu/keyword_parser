import docx
import argparse
import os
import re
from InvalidFileTypeError import InvalidFileTypeError

ORIGINAL_DOCX_FILES_DIRECTORY = 'original_docx_files'
UPDATED_DOCX_FILES_DIRECTORY = 'updated_docx_files'


def parse_file_name():
    parser = argparse.ArgumentParser(
        prog='keyword Parser',
        description="Finds and updates keywords wrapped with '[]' with user input")

    parser.add_argument("file_name", help='Name of the file to be parsed')
    file_name = parser.parse_args().file_name

    current_path = os.path.dirname(os.path.realpath(__file__))
    directory_path = os.path.join(current_path, ORIGINAL_DOCX_FILES_DIRECTORY)
    file_path = os.path.join(directory_path, file_name)

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f'docx file does not exist in the directory {directory_path}')

    _, file_extension = os.path.splitext(file_path)

    if file_extension.lower() != '.docx':
        raise InvalidFileTypeError(file_extension, '.docx')

    return file_path


def validate_directories():
    current_path = os.path.dirname(os.path.realpath(__file__))
    original_docx_files_path = os.path.join(current_path, 'original_docx_files')
    updated_docx_files_path = os.path.join(current_path, 'updated_docx_files')

    if not os.path.isdir(original_docx_files_path):
        raise NotADirectoryError(f"original_docx_files directory is not found within {current_path}")

    if not os.path.isdir(updated_docx_files_path):
        raise NotADirectoryError(f"updated_docx_files directory is not found within {current_path}")


def extract_keywords(document):
    keywords = set()
    pattern = re.compile(r'\[(\w+)]')

    for para in document.paragraphs:
        keywords.update(pattern.findall(para.text))

    return keywords


def main():
    validate_directories()
    file_path = parse_file_name()
    document = docx.Document(file_path)
    keywords = extract_keywords(document)


if __name__ == "__main__":
    main()
