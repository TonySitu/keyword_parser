import docx
import argparse


def parse_file_name(file_name):
    parser = argparse.ArgumentParser(
        prog='keyword Parser',
        description="Finds and updates keywords wrapped with '[]' with user input")
    parser.add_argument(file_name, help='Name of the file to be parsed')
    file_name = parser.parse_args()

    return file_name


def main():
    pass


if __name__ == "__main__":
    main()
