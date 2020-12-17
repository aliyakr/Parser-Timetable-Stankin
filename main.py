from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import json


# Таблицы XLSX (из файла)
wb = load_workbook(filename = 'timetable.xlsx')

with open('template.json', 'r') as f:
    jsn = json.load(f)

# Получаем название группы (расписания)
name_group = wb['Table 1']['A1'].value
name_group = ''.join(name_group.splitlines())
jsn['name_group'] = name_group


# Получаем первый столбец (time = 1)
test = str(wb['Table 2']['B2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][0] = test
print()

test = str(wb['Table 2']['B3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][1] = test

test = str(wb['Table 2']['B4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][2] = test

test = str(wb['Table 2']['B5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][3] = test

test = str(wb['Table 2']['B6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][4] = test

test = str(wb['Table 2']['B7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][0][5] = test


# Получаем первый столбец (time = 2)
test = str(wb['Table 2']['C2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][0] = test
print()

test = str(wb['Table 2']['C3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][1] = test

test = str(wb['Table 2']['C4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][2] = test

test = str(wb['Table 2']['C5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][3] = test

test = str(wb['Table 2']['C6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][4] = test

test = str(wb['Table 2']['C7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][1][5] = test


# Получаем первый столбец (time = 3)
test = str(wb['Table 2']['D2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][0] = test
print()

test = str(wb['Table 2']['D3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][1] = test

test = str(wb['Table 2']['D4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][2] = test

test = str(wb['Table 2']['D5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][3] = test

test = str(wb['Table 2']['D6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][4] = test

test = str(wb['Table 2']['D7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][2][5] = test


# Получаем первый столбец (time = 4)
test = str(wb['Table 2']['E2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][0] = test

test = str(wb['Table 2']['E3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][1] = test

test = str(wb['Table 2']['E4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][2] = test

test = str(wb['Table 2']['E5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][3] = test

test = str(wb['Table 2']['E6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][4] = test

test = str(wb['Table 2']['E7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][3][5] = test



# Получаем первый столбец (time = 5)
test = str(wb['Table 2']['F2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][0] = test

test = str(wb['Table 2']['F3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][1] = test

test = str(wb['Table 2']['F4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][2] = test

test = str(wb['Table 2']['F5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][3] = test

test = str(wb['Table 2']['F6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][4] = test

test = str(wb['Table 2']['F7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][4][5] = test



# Получаем первый столбец (time = 6)
test = str(wb['Table 2']['G2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][0] = test

test = str(wb['Table 2']['G3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][1] = test

test = str(wb['Table 2']['G4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][2] = test

test = str(wb['Table 2']['G5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][3] = test

test = str(wb['Table 2']['G6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][4] = test

test = str(wb['Table 2']['G7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][5][5] = test


# Получаем первый столбец (time = 7)
test = str(wb['Table 2']['H2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][0] = test

test = str(wb['Table 2']['H3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][1] = test

test = str(wb['Table 2']['H4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][2] = test

test = str(wb['Table 2']['H5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][3] = test

test = str(wb['Table 2']['H6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][4] = test

test = str(wb['Table 2']['H7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][6][5] = test


# Получаем первый столбец (time = 8)
test = str(wb['Table 2']['I2'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][0] = test

test = str(wb['Table 2']['I3'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][1] = test

test = str(wb['Table 2']['I4'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][2] = test

test = str(wb['Table 2']['I5'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][3] = test

test = str(wb['Table 2']['I6'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][4] = test

test = str(wb['Table 2']['I7'].value)
test = ''.join(test.splitlines())
jsn['array_lesson'][7][5] = test


with open("data_file.json", "w", encoding='utf-8') as write_file:
    json.dump(jsn, write_file, ensure_ascii=False)