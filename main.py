from cac import CAC
from uac import UAC
from utw import UniversityTW
from ceec import CEEC
from excel import write_excel
from progress import Progress
import os

def line():
    default_terminal_columns = os.get_terminal_size().columns
    print('\033[38;2;75;75;75m' + '─' * default_terminal_columns + '\033[0m')

year = int(input('請輸入\033[36m年份\033[0m：'))
line()

scores = []
for subject in ['國文', '英文', '數學A', '數學B', '社會', '自然']:
    score = int(input(f'請輸入\033[36m{subject}\033[0m級分：'))
    scores.append(score)
score = input(f'請輸入\033[36m英聽\033[0m成績：')
scores.append(score)
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
print('請輸入\033[36m學群編號\033[0m（以空格隔開）')
group_indexes = input('').split(' ')
group_indexes = [int(index) - 1 for index in group_indexes]
groups = list(groups.items())
groups = [groups[index] for index in group_indexes]
line()

departments = {}
for group_name, group_code in groups:
    print(f'正在取得\033[32m{group_name}\033[0m校系列表...')
    group_departments = cac.get_departments_by_group(group_code)
    for school_name, school_departments in group_departments.items():
        group_schools = departments.get(school_name, {})
        group_schools.update(school_departments)
        departments[school_name] = group_schools

# print(f'正在整理校系...')
departments = [
    (school_name, department_name, department_id)
    for school_name, school_departments in departments.items()
    for department_name, department_id in school_departments.items()
]
line()

print(f'正在取得校系簡章...')
infos = {
    department_id: cac.get_info(department_id)
    for school_name, department_name, department_id in Progress(departments)
}

print(f'正在取得{year-1}年分發入學錄取標準...')
uac = UAC(year - 1)
averages = {
    department_id: uac.get_average(school_name, department_name)
    for school_name, department_name, department_id in departments
}
school_averages = {}
for school_name, department_name, department_id in departments:
    max_average = school_averages.get(school_name, 0)
    average = averages[department_id]
    if average == '':
        average = 0
    max_average = max(max_average, average)
    school_averages[school_name] = max_average

print(f'正在取得個人申請篩選結果...')
utw = UniversityTW(year - 1)
screen_result_year = utw.support_year()
print(f'\033[31m● 目前僅支援{screen_result_year}年篩選結果\033[0m')
screen_results = {
    department_id: utw.get_screen_result(department_id)
    for school_name, department_name, department_id in Progress(departments)
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

# print(f'正在統整資料...')
excel = [
    (
        school_name,
        department_name,
        department_id,
        infos[department_id],
        averages[department_id],
        screen_results[department_id]
    )
    for school_name, department_name, department_id in departments
]
excel.sort(key=lambda x: (-school_averages[x[0]], x[2]))
# line()

print('正在寫入檔案...')
filename = f'{year}{"+".join(group_name[:-2] for group_name, group_code in groups)}學群學測標準.xlsx'
write_excel(
    year,
    screen_result_year,
    scores,
    excel,
    new_standards,
    old_standards,
    filename
)
print(f'已將結果寫入\033[32m{filename}\033[0m')
line()
input('請按\033[36mEnter鍵[↵]\033[0m繼續...')