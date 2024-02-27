import requests
from bs4 import BeautifulSoup
import xlrd

class CEEC:
    __slots__ = []
    def __init__(self):
        pass
    def get_standards(self, year):
        subjects = ['國文', '英文', '數學A', '數學B', '社會', '自然']
        standards = {
            subject: []
            for subject in subjects
        }

        url = 'https://www.ceec.edu.tw/xmdoc?xsmsid=0J018604485538810196'
        soup = self.request(url)
        a = soup.find(lambda tag: tag.name == 'a' and str(year) in tag.text)
        if a is None:
            raise ValueError
        
        url = f'https://www.ceec.edu.tw{a["href"]}'
        soup = self.request(url)
        a = soup.find(lambda tag: '各科成績標準一覽表' in tag.text and tag.get('href', '').endswith('xls'))
        
        url = a['href']
        response = requests.get(url)
        wb = xlrd.open_workbook(file_contents=response.content)
        sheet = wb.sheet_by_index(0)
        first_row = sheet.row_values(1)
        first_col = sheet.col_values(0)
        if '頂標' in first_row:
            subject_index = {
                subject: first_col.index(subject)
                for subject in subjects
            }

            index = first_row.index('頂標')
            for i in range(5):
                col = sheet.col_values(index+i)
                for subject in subjects:
                    standard = int(col[subject_index[subject]])
                    standards[subject].append(standard)
        else:
            subject_index = {
                subject: first_row.index(subject)
                for subject in subjects
            }

            index = first_col.index(year)
            for i in range(5):
                row = sheet.row_values(index+i)
                for subject in subjects:
                    standard = int(row[subject_index[subject]])
                    standards[subject].append(standard)
        return standards
    def request(self, url):
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup