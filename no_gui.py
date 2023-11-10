import os

from json_parser import convert

if __name__ == "__main__":
    current_dir = os.getcwd()
    for file in os.listdir(current_dir):
        if file.endswith('.json'):
            convert(file)