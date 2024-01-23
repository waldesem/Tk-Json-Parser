import json
import os
import shutil
import sys

import openpyxl
from tkinter import Tk, Button
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename


anketa_path = os.path.join(sys._MEIPASS, 'anketa.xlsx') \
    if getattr(sys, 'frozen', False) else 'anketa.xlsx'

def upload():
    file = askopenfilename(filetypes=[("Json files", ".json")])
    if file:
        fullname = convert(file)
        showinfo(title='Окончание операции', 
                 message=f"Импорт анкеты {fullname} завершен")
        root.destroy()


def convert(file):
    wb = openpyxl.load_workbook(anketa_path)
    sheet = wb.worksheets[0]

    with open(file, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
        fullname = []
        for key, value in data.items():
            match key:
                case 'positionName':
                    sheet['C3'] = value
                case 'department':
                    sheet['D3'] = value
                case 'statusDate':
                    sheet['A3'] = value
                case 'lastName':
                    fullname.append(value)
                case 'firstName':
                    fullname.append(value)
                case 'midName':
                    fullname.append(value)
                case 'birthday':
                    birthday = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
                    sheet['L3'] = birthday
                case 'birthplace':
                    sheet['M3'] = value
                case 'citizen':
                    sheet['T3'] = value
                case 'regAddress':
                    sheet['N3'] = value
                case 'validAddress':
                    sheet['O3'] = value
                case 'contactPhone':
                    sheet['Y3'] = value
                case 'email':
                    sheet['Z3'] = value
                case 'inn':
                    sheet['V3'] = value
                case 'snils':
                    sheet['U3'] = value
                case 'passportSerial':
                    sheet['P3'] = value
                case 'passportNumber':
                    sheet['Q3'] = value
                case 'passportIssueDate':
                    issue = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
                    sheet['R3'] = issue
                case 'nameWasChanged':
                    if isinstance(value, list):
                        previous = []
                        for item in value:
                            prev = [str(v) for k, v in item.items() \
                                    if k in ['yearOfChange', 'firstNameBeforeChange', 
                                            'lastNameBeforeChange', 
                                            'midNameBeforeChange', 'reason']]
                            previous.append(', '.join(prev))
                        sheet['S3'] = '; '.join(previous)
                case 'education':
                    if isinstance(value, list):
                        education = []
                        for item in value:
                            edu = [str(v) for k, v in item.items() \
                                    if k in ['endYear', 'institutionName', 
                                            'specialty']]
                            education.append(', '.join(edu))
                        sheet['X3'] = '; '.join(education)
                case 'experience':
                    if isinstance(value, list):
                        for index, item in enumerate(value[:3]):
                            for k, v in item.items():
                                match k:
                                    case 'name':
                                        sheet[f'AB{index + 3}'] = v
                                    case 'address':
                                        sheet[f'AC{index + 3}'] = v
                                    case 'position':
                                        sheet[f'AD{index + 3}'] = v
                                    case 'fireReason':
                                        sheet[f'AE{index + 3}'] = v
                            
                            period = ((
                                f"{item['beginDate'][-2:]}."
                                f"{item['beginDate'][5:7]}."
                                f"{item['beginDate'][:4]}"
                                ) if 'beginDate' in item else '') +  ' - ' + ((
                                f"{item['endDate'][-2:]}."
                                f"{item['endDate'][5:7]}."
                                f"{item['endDate'][:4]}"
                                ) if 'endDate' in item else '')
                            
                            sheet[f'AA{index + 3}'] = period

        full_name = ' '.join(fullname).rstrip()
        sheet['K3'] = full_name

    new_file = file.replace('.json', '.xlsx')
    wb.save(new_file)
    wb.close()

    os.startfile(new_file)
    
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
