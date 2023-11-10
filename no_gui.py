import os

from parser_json import Parser

if __name__ == "__main__":
    current_dir = os.getcwd()
    for file in os.listdir(current_dir):
        if file.endswith('.json'):
            Parser(file)