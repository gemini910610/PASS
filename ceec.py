import requests
from bs4 import BeautifulSoup
import pandas

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
        df = pandas.read_excel(url)

        subjects = ['國文', '英文', '數學A', '數學B', '社會', '自然']
        cols = df.iloc[0].tolist()
        rename_index = {
            '年度': 0,
            **{
                subject: cols.index(subject)
                for subject in subjects
            }
        }

        cols = df.columns
        rename_cols = {
            cols[index]: name
            for name, index in rename_index.items()
        }
        df = df.rename(columns=rename_cols)

        index = df[df['年度'] == year].index[0]
        standards = {
            subject: df[index:index+5][subject].tolist()
            for subject in subjects
        }

        return standards
    def request(self, url):
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup