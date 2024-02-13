from cac import CAC
from uac import UAC
from utw import UniversityTW
from ceec import CEEC
from excel import write_excel
import os
from tqdm.rich import tqdm

from tqdm import TqdmExperimentalWarning
import warnings
warnings.filterwarnings('ignore', category=TqdmExperimentalWarning)

def line():
    print('\033[38;2;75;75;75m' + '━' * os.get_terminal_size().columns + '\033[0m')

os.system('cls')
year = int(input('請輸入\033[36m年份\033[0m：'))
line()

scores = []
for subject in ['國文', '英文', '數學A', '數學B', '社會', '自然']:
    score = int(input(f'請輸入\033[36m{subject}\033[0m級分：'))
    scores.append(score)
score = input(f'請輸入\033[36m英聽\033[0m成績：')
line()

cac = CAC(year)
print('正在取得學群列表...')
groups = cac.get_groups()
line()
width = max(len(group) for group in groups)
for i, group in enumerate(groups):
    end = '\n' if (i+1) % 4 == 0 else '    '
    length = width - len(group)
    print(f'{i+1:>2}. {group}{"  "*length}', end=end)
print()
group_index = int(input('請輸入\033[36m學群編號\033[0m：')) - 1
group_name, group_code = list(groups.items())[group_index]
line()

print(f'正在取得\033[32m{group_name}\033[0m校系列表...')
departments = cac.get_departments_by_group(group_code)
line()

print(f'正在取得校系簡章...')
infos = {
    department_id: cac.get_info(department_id)
    for school_name, school_departments in tqdm(departments.items())
    for department_name, department_id in school_departments.items()
}

print(f'正在取得{year-1}年分發入學錄取標準...')
uac = UAC(year - 1)
averages = {
    department_id: uac.get_average(school_name, department_name)
    for school_name, school_departments in departments.items()
    for department_name, department_id in school_departments.items()
}

print(f'正在取得個人申請篩選結果...')
utw = UniversityTW(year - 1)
screen_result_year = utw.support_year()
print(f'\033[31m● 目前僅支援{screen_result_year}年篩選結果\033[0m')
screen_results = {
    department_id: utw.get_screen_result(department_id)
    for school_name, school_departments in tqdm(departments.items())
    for department_name, department_id in school_departments.items()
}

print('正在取得各科成績標準...')
ceec = CEEC()
old_standards = ceec.get_standards(year - 1)
try:
    new_standards = ceec.get_standards(year)
except ValueError:
    print(f'\033[31m● 「{year}各科成績標準一覽表」尚未公布，使用{year-1}年資料作為參考\033[0m')
    new_standards = old_standards
line()

print(f'正在統整資料...')
excel = [
    (
        school_name,
        department_name,
        department_id,
        infos[department_id],
        averages[department_id],
        screen_results[department_id]
    )
    for school_name, school_departments in departments.items()
    for department_name, department_id in school_departments.items()
]
line()

print('正在寫入檔案...')
write_excel(
    year,
    screen_result_year,
    scores,
    excel,
    new_standards,
    old_standards,
    f'{year}{group_name}學測標準.xlsx'
)
print(f'已將結果寫入\033[32m{year}{group_name}學測標準.xlsx\033[0m')
line()
os.system('pause')