import requests
from bs4 import BeautifulSoup
import xlrd

class CEEC:
    __slots__ = []
    def __init__(self):
        pass
    def get_standards(self, year):
        url = 'https://www.ceec.edu.tw/xmdoc?xsmsid=0J018604485538810196'
        soup = self.request(url)
        a = soup.find(lambda tag: tag.name == 'a' and str(year) in tag.text)
        if a is None:
            raise ValueError
        
        url = f'https://www.ceec.edu.tw{a["href"]}'
        soup = self.request(url)
        a = soup.find(string='各科成績標準一覽表').parent
        
        url = a['href']
        response = requests.get(url)
        wb = xlrd.open_workbook(file_contents=response.content)
        sheet = wb.sheet_by_index(0)
        row = sheet.row_values(1)
        subjects = ['國文', '英文', '數學A', '數學B', '社會', '自然']
        subject_index = {
            subject: row.index(subject)
            for subject in subjects
        }

        index = sheet.col_values(0).index(year)
        standards = {
            subject: []
            for subject in subjects
        }
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