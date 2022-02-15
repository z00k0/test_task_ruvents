import openpyxl
import datetime
import calendar


file = 'task_support.xlsx'
wb = openpyxl.load_workbook(file)
sheet_name = wb.sheetnames
sheet = wb['Tasks']


def test1():
    test1_list = []
    for cell in sheet['B3':'B1002']:
        test1_list.append(int(cell[0].value))

    result = 0
    for num in test1_list:
        if num % 2 == 0:
            result += 1
    return result


print(f'Количество четных чисел в столбце В: {test1()}')


def test2():
    def is_prime(num):
        start = 2
        end = int(num ** .5) + 1
        for i in range(start, end):
            if num % i == 0:
                return False
        return True

    test2_list = []
    for cell in sheet['C3':'C1002']:
        test2_list.append(int(cell[0].value))

    count_of_primes = 0
    for num in test2_list:
        if is_prime(num):
            count_of_primes += 1
    return count_of_primes


print(f'Количество простых чисел в столбце С: {test2()}')


def test3():
    test3_list = []
    for cell in sheet['D3':'D1002']:
        val = cell[0].value.replace(' ', '').replace(',', '.')
        test3_list.append(float(val))

    result = 0
    for num in test3_list:
        if num < .5:
            result += 1
    return result


print(f'Количество чисел меньше 0.5 в столбце D: {test3()}')


def test4():
    test4_list = []
    for cell in sheet['E3':'E1002']:
        val = cell[0].value.split()[0]
        if val == 'Tue':
            test4_list.append(val)
    return len(test4_list)


print(f'Количество вторников в столбце E: {test4()}')


def test5():
    result = 0
    for cell in sheet['F3':'F1002']:
        val = cell[0].value.split()[0]  # сплитом разделяю дату и время, индекс 0 отбирает только дату
        year, month, day = val.split('-')
        if datetime.date(int(year), int(month), int(day)).weekday() == 1:  # проверка, является ли дата вторником
            result += 1
    return result


print(f'Количество вторников в столбце F: {test5()}')


def test6():
    result = 0
    for cell in sheet['G3':'G1002']:
        val = cell[0].value
        month, day, year = val.split('-')
        cell_date = datetime.date(int(year), int(month), int(day))
        last_day_of_month = calendar.monthrange(int(year), int(month))[1]
        last_7_days_of_month = datetime.date(int(year), int(month), last_day_of_month) - datetime.timedelta(days=7)
        """
        Чтобы определить, является ли вторник последним в текущем месяце вычисляется последний день месяца
        last_day_of_month
        Из последнего дня месяца вычитаю 7 дней. Если дата в ячейке больше чем last_7_days_of_month, 
        и является вторником, инкрементирую счетчик
        """
        if cell_date.weekday() == 1 and cell_date > last_7_days_of_month:
            # print(f'{year}-{month}-{day} ')
            result += 1
    return result


print(f'Количество последних вторников в столбце G: {test6()}')
