# Small JSON to XLSX Importer

## Project Description
This project provides a Python script for importing data from a JSON file to an XLSX format. 
It uses the `openpyxl` library to create an XLSX file, the `json` library to parse the JSON data, and the `tkinter` library for the Tkinter GUI.
Python 3.10+ is required.

## Installation
To use this script, follow these steps:
Clone the repository: 
```
git clone https://github.com/waldesem/Json_Excel.git
````
Install the required dependencies: 
```
pip install -r requirements.txt
````
For building executables: 
```
pyinstaller --clean json_parser.py
```
Add file 'anketa.xlsx' to './dist/tk_gui' folder 
```
cp anketa.xlsx ./dist/json_parser
```

## Usage
1. Run the script: `python json_parser.py`
2. Run executables: `./dist/json_parser`

## License
This project is licensed under the MIT License.