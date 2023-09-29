
from openpyxl import load_workbook
from .config import CON
#------------------------------------------------------------------------------------------

fn = 'Analytics_sending_app.xlsx'
wb = load_workbook(fn)
ws = wb.active

cod_MPK=[]
item = 2
while ws[f'D{item}'].value is not None: item += 1
index_last_MPK = item

for i in range(2, index_last_MPK):
    cod_MPK.append(ws[f'D{i}'].value)

cod_MPK_sort = []
for i in cod_MPK:
    if i[:4:] in CON:
        cod_MPK_sort.append(i)
    else:
        cod_MPK_sort.append(i[:4:])

cod_MPK_count = {}
for i in cod_MPK_sort:
    cod_MPK_count[f'{i}'] = cod_MPK_sort.count(i)

cod_MPK_count_sort = {}
for key, value in cod_MPK_count.items():
    if key not in cod_MPK_count_sort:
        cod_MPK_count_sort[key] = value

# for key, value in cod_MPK_count_sort.items():
#     print(f'  {key} : {value}')
#     print('\n')
#     sum_ += value
# print(f'    колличество заявкок - {sum_}')
#----------------------------------------------------------------------------------------------------------------

ws = wb['analytics data']

for key, value in cod_MPK_count_sort.items():
    if key[:4:] in CON:
        continue
    for i in range(5, 776):
        if ws[f'D{i}'].value is not None:
            a = ws[f'D{i}'].value
            b = ws[f'E{i}'].value
            if key == a:
                if b is not None: ws[f'E{i}'] = f'{int(b) + int(value)}'
                ws[f'E{i}'] = value

wb.save(fn)
wb.close()

sum_ = 0
for key, value in cod_MPK_count_sort.items():
    if len(key) > 4:
        print(f'  {key} : {value}')
        print('\n')
    sum_ += value
print(f'    колличество заявок - {sum_}')