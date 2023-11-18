import json

import openpyxl
from tkinter import Tk, Button
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename


def upload():
    file = askopenfilename(filetypes=[("Json files", ".json")])

    convert(file)
    
    showinfo(title='Окончание операции', message='Конвертация завершена')

    root.destroy()

def convert(file):
    wb = openpyxl.load_workbook('anketa.xlsx', keep_vba=True)
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
                    sheet['L3'] = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
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
                    sheet['R3'] = f"{value[-2:]}.{value[5:7]}.{value[:4]}"
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
                        for index, item in enumerate(value):
                            for k, v in item.items():
                                match k:
                                    case 'workplace':
                                        sheet[f'AB{index + 3}'] = v
                                    case 'address':
                                        sheet[f'AC{index + 3}'] = v
                                    case 'position':
                                        sheet[f'AD{index + 3}'] = v
                                    case 'fireReason':
                                        sheet[f'AE{index + 3}'] = v
                            
                            sheet[f'AA{index + 3}'] = ((
                                f"{item['beginDate'][-2:]}."
                                f"{item['beginDate'][5:7]}."
                                f"{item['beginDate'][:4]}"
                                ) if 'beginDate' in item else '') +  ' - ' + ((
                                f"{item['endDate'][-2:]}."
                                f"{item['endDate'][5:7]}."
                                f"{item['endDate'][:4]}"
                                ) if 'endDate' in item else '')
                            
        sheet['K3'] = ' '.join(fullname).rstrip()

    wb.save(file.replace('json', 'xlsx'))


if __name__ == '__main__':
    root = Tk()
    root.title('JSON Parser')
    Button(root, text='Загрузить JSON', command=upload).pack(padx=60, pady=50)
    root.mainloop()
