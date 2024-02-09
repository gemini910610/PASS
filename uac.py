import requests
from io import BytesIO
import pdfplumber
import pandas
from pandas import DataFrame

class UAC:
    def __init__(self, year):
        self.year = year
        url = f'https://www.uac.edu.tw/{year}data/{year}_result_school_data.pdf'
        response = requests.get(url)
        io = BytesIO(response.content)
        pdf = pdfplumber.open(io)
        tables = [
            DataFrame(page.extract_table())
            for page in pdf.pages
        ]
        for table in tables:
            table.columns = table.iloc[0].tolist()
            table.drop(0, inplace=True)
        table = pandas.concat(tables)

        schools = table['校名'].tolist()
        schools = [
            school.replace('\n', '')
            for school in schools
        ]
        table['校名'] = schools

        departments = table['系組名'].tolist()
        departments = [
            department.replace('\n', '')
            for department in departments
        ]
        table['系組名'] = departments

        weights = table['採計及加權'].tolist()
        weights = [
            sum(
                float(rate.split('x')[-1])
                for rate in weight.split(' ')
            )
            for weight in weights
        ]
        scores = table['普通生\n錄取分數'].tolist()
        scores = [
            0 if score == '-----' else float(score)
            for score in scores
        ]
        averages = [
            round(score / weight, 2)
            for score, weight in zip(scores, weights)
        ]
        table['平均'] = averages

        table = table[['校名', '系組名', '平均']].copy()
        self.table = table
    def get_average(self, school_name, department_name):
        mask = (self.table['校名'] == school_name) & (self.table['系組名'] == department_name)
        result = self.table[mask]
        if result.empty:
            return ''
        return result['平均'].tolist()[0]