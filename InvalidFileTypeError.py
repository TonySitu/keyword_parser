
class InvalidFileTypeError(Exception):
    def __init__(self, file_type, expected_type):
        self.file_type = file_type
        self.expected_type = expected_type
        super().__init__(f"Invalid file type: {file_type}. Expected a {expected_type} file.")
