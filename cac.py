import requests
from bs4 import BeautifulSoup
from tqdm.rich import tqdm

class CAC:
    def __init__(self, year):
        self.year = year
        url = f'https://www.cac.edu.tw/apply{year}/query.php'
        soup = self.request('get', url)
        tag = soup.find(string='查詢所有校系').parent
        href = tag['href']
        index = href.rfind('/')
        self.host = f'https://www.cac.edu.tw/apply{year}/{href[:index]}'
    def request(self, method, url, data=None):
        headers = {'User-Agent': 'Chrome/121.0.0.0'}
        response =requests.request(method, url, headers=headers, data=data)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup
    # def get_schools(self):
    #     url = f'{self.host}/TotalGsdShow.htm'
    #     soup = self.request('get', url)
    #     tags = soup.select('a')
    #     schools = {}
    #     for a in tags:
    #         text = a.text.split(' ')[0]
    #         id = text[1:4]
    #         name = text[5:]
    #         schools[name] = id
    #     return schools
    # def get_departments_by_school(self, school_id):
    #     url = f'{self.host}/ShowGsd.php'
    #     data = {
    #         'TxtGsdCode': school_id,
    #         'SubTxtGsdCode': '依校系代碼查詢',
    #         'action': 'SubTxtGsdCode'
    #     }
    #     soup = self.request('post', url, data)
    #     tags = soup.select('b')
    #     departments = {}
    #     for b in tags:
    #         texts = [
    #             text
    #             for text in b.text.split(' ')
    #             if text != ''
    #         ]
    #         id = texts[-1][1:-1]
    #         school = texts[0]
    #         department = ''.join(texts[1:-1])
    #         school_departments = departments.get(school, {})
    #         school_departments[department] = id
    #         departments[school] = school_departments
    #     return departments
    def get_groups(self):
        url = f'{self.host}/findgsdgroup.htm'
        soup = self.request('get', url)
        tags = soup.select('option')
        groups = {}
        for option in tags:
            group = option.text.strip()
            code = option['value']
            groups[group] = code
        return groups
    def get_departments_by_group(self, group_code):
        url = f'{self.host}/findgsdname.php'
        data = {'gcode': group_code}
        soup = self.request('post', url, data)
        tags = soup.select('option')
        group_departments = [
            option.text.strip()
            for option in tags
        ]

        url = f'{self.host}/findgsd.php'
        departments = {}
        for group in tqdm(group_departments):
            data = {
                'GsdName': group,
                'gcode': group_code
            }
            soup = self.request('post', url, data)
            tags = soup.select('b')
            for b in tags:
                texts = [
                    text
                    for text in b.text.split(' ')
                    if text != ''
                ]
                id = texts[-1][1:-1]
                school = texts[0]
                department = ''.join(texts[1:-1])
                school_departments = departments.get(school, {})
                school_departments[department] = id
                departments[school] = school_departments
        return departments
    def get_info(self, department_id):
        url = f'{self.host}/html/{self.year}_{department_id}.htm'
        soup = self.request('get', url)
        tags = soup.select('tr')

        tr = tags[3]
        tds = tr.select('td')
        td = tds[2]
        rename_subjects = {
            '國文': '國',
            '英文': '英',
            '數學A': '數A',
            '數學B': '數B',
            '社會': '社',
            '自然': '自',
            '英聽': '英聽'
        }
        subjects = [
            rename_subjects.get(
                BeautifulSoup(text, 'html.parser').text,
                '組合'
            )
            for text in str(td).split('<br/>')
        ]
        td = tds[3]
        standards = [
            BeautifulSoup(text, 'html.parser').text.replace('級', '')
            for text in str(td).split('<br/>')
        ]
        td = tds[4]
        rates = [
            BeautifulSoup(text, 'html.parser').text
            for text in str(td).split('<br/>')
        ]
        subjects = {
            subject: (standard, rate)
            for subject, standard, rate in zip(subjects, standards, rates)
        }
        td = tds[7]
        meetings = [
            BeautifulSoup(text, 'html.parser').text
            for text in str(td).split('<br/>')
        ]
        meetings.remove('審查資料')

        tr = tags[4]
        td = tr.select('td')[-1]
        number = td.text

        tr = tags[11]
        td = tr.select('td')[-1]
        send_time = td.text

        tr = tags[12]
        td = tr.select('td')[-1]
        end_time = td.text

        tr = tags[13]
        td = tr.select('td')[1]
        meeting_time = ' '.join([
            BeautifulSoup(text, 'html.parser').text
            for text in str(td).split('<br/>')
        ])
        if meeting_time == '--':
            meeting_time = ''

        return {
            'subjects': subjects,
            'meetings': meetings,
            'number': number,
            'send_time': send_time,
            'end_time': end_time,
            'meeting_time': meeting_time
        }