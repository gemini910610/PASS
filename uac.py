import requests
from io import BytesIO
import pdfplumber
from progress import Progress

class UAC:
    __slots__ = ['year', 'averages']
    def __init__(self, year):
        self.year = year
        url = f'https://www.uac.edu.tw/{year}data/{year}_result_school_data.pdf'
        response = requests.get(url)
        io = BytesIO(response.content)
        pdf = pdfplumber.open(io)

        table = []
        header = None
        for page in Progress(pdf.pages):
            rows = page.extract_table()
            if header is None:
                header = rows[0]
            table.extend(rows[1:])
        column_index = {
            title: header.index(title)
            for title in ['校名', '系組名', '採計及加權', '普通生\n錄取分數']
        }

        averages = {}
        for row in table:
            score = row[column_index['普通生\n錄取分數']]
            if score == '-----':
                continue
            school_name = row[column_index['校名']].replace('\n', '').replace(' ', '')
            department_name = row[column_index['系組名']].replace('\n', '').replace(' ', '')
            weight = sum(
                float(rate.split('x')[-1])
                for rate in row[column_index['採計及加權']].split(' ')
            )
            score = float(score)
            average = round(score / weight, 2)

            department_averages = averages.get(school_name, {})
            department_averages[department_name] = average
            averages[school_name] = department_averages
        self.averages = averages
    def get_average(self, school_name, department_name):
        try:
            return self.averages[school_name][department_name]
        except:
            return ''