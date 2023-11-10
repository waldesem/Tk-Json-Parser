import json

import openpyxl


class Parser:
    """ Create class for import data from json file in xlsx format"""

    def __init__(self, file) -> None:
        with open(file, 'r', newline='', encoding='utf-8-sig') as f:
            self.json_dict = json.load(f)

            self.wb = openpyxl.load_workbook('anketa.xlsx', keep_vba=True)
            self.sheet = self.wb.worksheets[0]

            self.sheet['A3'] = self.json_dict['statusDate']
            self.sheet['C3'] = self.json_dict['positionName'].strip() \
                    if 'positionName' in self.json_dict else ''
            self.sheet['D3'] = self.json_dict['department'].strip() \
                    if 'department' in self.json_dict else ''
            self.sheet['K3'] = self.parse_fullname()
            self.sheet['L3'] = self.parse_date(self.json_dict['birthday'])\
                    if 'birthday' in self.json_dict \
                        else ''
            self.sheet['M3'] = self.json_dict['birthplace'].strip() \
                    if 'birthplace' in self.json_dict else ''
            self.sheet['N3'] = self.json_dict['regAddress'].strip() \
                    if 'regAddress' in self.json_dict else ''
            self.sheet['O3'] = self.json_dict['validAddress'].strip() \
                    if 'validAddress' in self.json_dict else ''
            self.sheet['P3'] = self.json_dict['passportSerial'].strip() \
                    if 'passportSerial' in self.json_dict else ''
            self.sheet['Q3'] = self.json_dict['passportNumber'].strip() \
                    if 'passportNumber' in self.json_dict else ''
            self.sheet['R3'] = self.parse_date(self.json_dict['passportIssueDate']) \
                    if 'passportIssueDate' in self.json_dict \
                        else ''
            self.sheet['S3'] = self.parse_previous()
            self.sheet['T3'] = self.json_dict['citizen'] \
                    if 'citizen' in self.json_dict else ''
            self.sheet['U3'] = self.json_dict['snils'].strip() \
                    if 'snils' in self.json_dict else ''
            self.sheet['V3'] = self.json_dict['inn'].strip() \
                    if 'inn' in self.json_dict else ''
            self.sheet['X3'] = self.parse_education()
            self.sheet['Y3'] = self.json_dict['contactPhone'].strip() \
                    if 'contactPhone' in self.json_dict else ''
            self.sheet['Z3'] = self.json_dict['email'].strip() \
                    if 'email' in self.json_dict else ''
            
            self.workplaces = self.parse_workplace()
            
            if len(self.workplaces):
                for index, value in enumerate(self.workplaces):
                    start_date = value['start_date']
                    end_date = value['end_date']
                    self.sheet[f'AA{index + 3}'] = f"{start_date} - {end_date}"
                    self.sheet[f'AB{index + 3}'] = value['workplace']
                    self.sheet[f'AC{index + 3}'] = value['address']
                    self.sheet[f'AD{index + 3}'] = value['position']
                    self.sheet[f'AE{index + 3}'] = value['reason']
        
        self.wb.save(file.replace('json', 'xlsx'))

    def parse_fullname(self):
        lastName = self.json_dict['lastName'].strip() \
            if 'midName' in self.json_dict else ''
        firstName = self.json_dict['firstName'].strip() \
            if 'midName' in self.json_dict else ''
        midName = self.json_dict['midName'].strip() \
            if 'midName' in self.json_dict else ''
        return f"{lastName} {firstName} {midName}".rstrip()
    
    def parse_date(self, data):
        new_data = (f"{data[-2:]}.{data[5:7]}.{data[:4]}")
        return new_data

    def parse_previous(self):
        if 'hasNameChanged' in self.json_dict:
            if len(self.json_dict['nameWasChanged']):
                previous = []
                for item in self.json_dict['nameWasChanged']:
                    firstNameBeforeChange = item['firstNameBeforeChange'].strip() \
                        if 'firstNameBeforeChange' in item else ''
                    lastNameBeforeChange = item['lastNameBeforeChange'].strip() \
                        if 'lastNameBeforeChange' in item else ''
                    midNameBeforeChange = item['midNameBeforeChange'].strip() \
                        if 'midNameBeforeChange' in item else ''
                    yearOfChange = str(item['yearOfChange']) \
                        if 'yearOfChange' in item else 'Дата отсутствует'
                    
                    previous.append(f"{yearOfChange} - {firstNameBeforeChange} "
                                    f"{lastNameBeforeChange} {midNameBeforeChange}".\
                                        rstrip())
                return '; '.join(previous)
        return ''
    
    def parse_education(self):
        if 'education' in self.json_dict:
            if len(self.json_dict['education']):
                education = []
                for item in self.json_dict['education']:
                    institutionName = item['institutionName'] \
                        if 'institutionName' in item else 'Нет данных'
                    beginYear = item['beginYear'] if 'specialty' in item else 'Неизвестно'
                    endYear = item['endYear'] if 'specialty' in item else 'н.в.'
                    specialty = item['specialty'] if 'specialty' in item else 'Неизвестно'

                    education.append(f"{str(beginYear)}-{str(endYear)} - "
                                     f"{institutionName}, {specialty}".strip())
                return '; '.join(education)
        return ''
    
    def parse_workplace(self):
        experience = []
        if self.json_dict['hasJob'] and 'experience' in self.json_dict:
            if len(self.json_dict['experience']):
                for item in self.json_dict['experience']:
                    work = {
                        'start_date': self.parse_date(item['beginDate']) \
                            if 'beginDate' in item else '',
                        'end_date': self.parse_date(item['endDate']) \
                            if 'endDate' in item else 'н.в.',
                        'workplace': item['name'].strip() if 'name' in item else '',
                        'address': item['address'].strip() if 'address' in item else '',
                        'position': item['position'].strip() if 'position' in item else '',
                        'reason': item['fireReason'].strip() if 'fireReason' in item else ''
                    }
                    experience.append(work)
        return experience
