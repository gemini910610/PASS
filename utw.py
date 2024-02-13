import requests
from bs4 import BeautifulSoup

class UniversityTW:
    __slots__ = ['year']
    def __init__(self, year):
        self.year = year
    def get_screen_result(self, department_id):
        school_id = department_id[:3]
        url = f'https://university-tw.ldkrsi.men/caac/{school_id}/{department_id}'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        tags = soup.select('.filter-result-table tr')
        if len(tags) == 0:
            return {}
        tds = tags[1].select('td')
        tds = tds[:-1]
        result = {}
        for td in tds:
            subject, score = td.text.split('=')
            result[subject] = score
        return result
    def support_year(self):
        url = 'https://university-tw.ldkrsi.men/caac/001/001012'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        tag = soup.find(lambda tag: tag.name == 'dt' and '篩選結果' in tag.text)
        year = int(tag.text[:3])
        return year