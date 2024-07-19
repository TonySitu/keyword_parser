from docx import Document
from docx.shared import Pt
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

    return original_docx_files_path, updated_docx_files_path


def extract_keywords(document):
    keywords = set()
    pattern = re.compile(r'\[(\w+)]')

    for para in document.paragraphs:
        keywords.update(pattern.findall(para.text))

    return keywords


def parse_replacements(keyword):
    print(f'What is the replacement for keyword: {keyword}?')
    return input('Enter the replacement: ')


def get_user_replacements(keywords):
    replacements = {}
    for keyword in keywords:
        replacement = parse_replacements(keyword)
        replacements[keyword] = replacement
    return replacements


def replace_keywords(text, replacements):
    pattern = re.compile(r'\[(\w+)]')

    def replacer(match):
        keyword = match.group(1)
        return replacements.get(keyword, match.group(0))

    return pattern.sub(replacer, text)


def process_doc(document, replacements):
    formatted_text = []
    for para in document.paragraphs:
        para_info = {
            'paragraph_style': para.style.name,
            'alignment': para.alignment,
            'runs': []}

        for run in para.runs:
            run_info = {
                'text': run.text,
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'font_name': run.font.name,
                'font_size': run.font.size.pt if run.font.size else None,
                'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
            }
            para_info['runs'].append(run_info)
        formatted_text.append(para_info)

    return formatted_text


def save_replaced_text_with_formatting(file_path, formatted_text):
    doc = Document()

    for para_info in formatted_text:
        para = doc.add_paragraph(style=para_info['paragraph_style'])
        para.alignment = para_info['alignment']

        for run_info in para_info['runs']:
            run = para.add_run(run_info['text'])
            if run_info['bold']:
                run.bold = run_info['bold']
            if run_info['italic']:
                run.italic = run_info['italic']
            if run_info['underline']:
                run.underline = run_info['underline']
            if run_info['font_name']:
                run.font.name = run_info['font']
            if run_info['font_size']:
                run.font.size = Pt(run_info['font_size'])
            if run_info['color']:
                run.font.color.rgb = run_info['color']

    doc.save(file_path)


def main():
    original_docx_files_path, updated_docx_files_path = validate_directories()
    file_path = 'C:\\Users\\Tonyn\\Desktop\\Projects\\keyword_parser\\original_docx_files\\CV.docx'
    document = Document(file_path)
    keywords = extract_keywords(document)
    replacements = get_user_replacements(keywords)
    formatted_text = process_doc(document, replacements)
    save_replaced_text_with_formatting(updated_docx_files_path + '\\lif.docx', formatted_text)


if __name__ == "__main__":
    main()
