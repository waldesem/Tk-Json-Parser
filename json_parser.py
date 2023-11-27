import json
import os
import shutil
import sys
import configparser

import openpyxl
from tkinter import Tk, Button
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename


anketa_path = os.path.join(sys._MEIPASS, 'anketa.xlsx') \
    if getattr(sys, 'frozen', False) else 'anketa.xlsx'

conclusion_path = os.path.join(sys._MEIPASS, 'conclusion.xlsm') \
    if getattr(sys, 'frozen', False) else 'conclusion.xlsm'

config_path = os.path.join(sys._MEIPASS, 'config.ini') \
    if getattr(sys, 'frozen', False) else 'config.ini'

config = configparser.ConfigParser()
config.read(config_path)
base_path = r'' + os.path.abspath(config.get('Settings', 'path'))


def upload():
    file = askopenfilename(filetypes=[("Json files", ".json")])
    if file:
        fullname = convert(file)
        showinfo(title='Окончание операции', 
                 message=f"Импорт анкеты {fullname} завершен")
        root.destroy()


def convert(file):
    wb_anketa = openpyxl.load_workbook(anketa_path)
    anketa_sheet = wb_anketa.worksheets[0]

    wb_conclusion = openpyxl.load_workbook(conclusion_path, keep_vba=True)
    conclusion_sheet = wb_conclusion.worksheets[0]

    with open(file, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
        fullname = []
        for key, value in data.items():
            match key:
                case 'positionName':
                    anketa_sheet['C3'] = value
                    conclusion_sheet['C4'] = value
                case 'department':
                    anketa_sheet['D3'] = value
                    conclusion_sheet['C5'] = value
                case 'statusDate':
                    anketa_sheet['A3'] = value
                case 'lastName':
                    fullname.append(value)
                case 'firstName':
                    fullname.append(value)
                case 'midName':
                    fullname.append(value)
                case 'birthday':
                    birthday = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
                    anketa_sheet['L3'] = birthday
                    conclusion_sheet['C8'] = birthday
                case 'birthplace':
                    anketa_sheet['M3'] = value
                case 'citizen':
                    anketa_sheet['T3'] = value
                case 'regAddress':
                    anketa_sheet['N3'] = value
                case 'validAddress':
                    anketa_sheet['O3'] = value
                case 'contactPhone':
                    anketa_sheet['Y3'] = value
                case 'email':
                    anketa_sheet['Z3'] = value
                case 'inn':
                    anketa_sheet['V3'] = value
                    conclusion_sheet['C10'] = value
                case 'snils':
                    anketa_sheet['U3'] = value
                case 'passportSerial':
                    anketa_sheet['P3'] = value
                    conclusion_sheet['C9'] = value          
                case 'passportNumber':
                    anketa_sheet['Q3'] = value
                    conclusion_sheet['D9'] = value
                case 'passportIssueDate':
                    issue = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
                    anketa_sheet['R3'] = issue
                    conclusion_sheet['E9'] = issue
                case 'nameWasChanged':
                    if isinstance(value, list):
                        previous = []
                        for item in value:
                            prev = [str(v) for k, v in item.items() \
                                    if k in ['yearOfChange', 'firstNameBeforeChange', 
                                            'lastNameBeforeChange', 
                                            'midNameBeforeChange', 'reason']]
                            previous.append(', '.join(prev))
                        anketa_sheet['S3'] = '; '.join(previous)
                        conclusion_sheet['C7'] = '; '.join(previous)
                case 'education':
                    if isinstance(value, list):
                        education = []
                        for item in value:
                            edu = [str(v) for k, v in item.items() \
                                    if k in ['endYear', 'institutionName', 
                                            'specialty']]
                            education.append(', '.join(edu))
                        anketa_sheet['X3'] = '; '.join(education)
                case 'experience':
                    if isinstance(value, list):
                        for index, item in enumerate(value[:3]):
                            for k, v in item.items():
                                match k:
                                    case 'name':
                                        anketa_sheet[f'AB{index + 3}'] = v
                                        conclusion_sheet[f'D{index + 11}'] = v
                                    case 'address':
                                        anketa_sheet[f'AC{index + 3}'] = v
                                    case 'position':
                                        anketa_sheet[f'AD{index + 3}'] = v
                                    case 'fireReason':
                                        anketa_sheet[f'AE{index + 3}'] = v
                            
                            period = ((
                                f"{item['beginDate'][-2:]}."
                                f"{item['beginDate'][5:7]}."
                                f"{item['beginDate'][:4]}"
                                ) if 'beginDate' in item else '') +  ' - ' + ((
                                f"{item['endDate'][-2:]}."
                                f"{item['endDate'][5:7]}."
                                f"{item['endDate'][:4]}"
                                ) if 'endDate' in item else '')
                            
                            anketa_sheet[f'AA{index + 3}'] = period
                            conclusion_sheet[f'C{index + 11}'] = period

        full_name = ' '.join(fullname).rstrip()
        anketa_sheet['K3'] = full_name
        conclusion_sheet['C6'] = full_name

    dir_name = make_folder(file, full_name)

    wb_anketa.save(os.path.join(dir_name, f'Анкета {full_name}.xlsx'))
    wb_anketa.close()
    wb_conclusion.save(os.path.join(dir_name, f'Заключение {full_name}.xlsm'))
    wb_conclusion.close()
    shutil.move(dir_name, os.path.join(base_path, full_name))
    
    return full_name


def make_folder(file, full_name):
    dir_name = os.path.join(os.path.dirname(file), full_name)
    if not os.path.isdir(dir_name):
        os.mkdir(dir_name)
        shutil.copyfile(file, os.path.join(dir_name, f'{full_name}.json'))
        return dir_name


if __name__ == '__main__':
    root = Tk()
    root.title('JSON Parser')
    root.resizable(False, False)
    Button(root, text='Загрузить JSON', command=upload).pack(padx=60, pady=50)
    root.mainloop()
