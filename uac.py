import requests
from io import BytesIO
import tabula
import pandas

class UAC:
    def __init__(self, year):
        url = f'https://www.uac.edu.tw/{year}data/{year}_result_school_data.pdf'
        response = requests.get(url)
        io = BytesIO(response.content)
        pages = tabula.read_pdf(io, pages='all')
        pdf = pandas.concat(pages)

        schools = pdf['校名'].tolist()
        schools = [
            school.replace('\r', '')
            for school in schools
        ]
        pdf['校名'] = schools

        departments = pdf['系組名'].tolist()
        departments = [
            department.replace('\r', '')
            for department in departments
        ]
        pdf['系組名'] = departments

        weights = pdf['採計及加權'].tolist()
        weights = [
            sum(
                float(rate.split('x')[-1])
                for rate in weight.split(' ')
            )
            for weight in weights
        ]
        scores = pdf['普通生\r錄取分數'].tolist()
        scores = [
            0 if score == '-----' else float(score)
            for score in scores
        ]
        averages = [
            round(score / weight, 2)
            for score, weight in zip(scores, weights)
        ]
        pdf['平均'] = averages

        pdf = pdf[['校名', '系組名', '平均']].copy()
        self.pdf = pdf
    def get_average(self, school_name, department_name):
        mask = (self.pdf['校名'] == school_name) & (self.pdf['系組名'] == department_name)
        result = self.pdf[mask]
        if result.empty:
            return ''
        return result['平均'].tolist()[0]