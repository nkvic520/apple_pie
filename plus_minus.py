from random import *

import openpyxl as op


def op_to_excel(data, file_name):
    wb = op.Workbook()
    ws = wb['Sheet']
    for i in range(4):
        for j in range(20):
            ws.cell(row=j + 1, column=i + 1).value = data[(20 * i) + j]
    wb.save(file_name)


def create_data_2(total):
    operation_type = randint(0, 1)
    if operation_type == 1:
        addend1 = randint(1, total - 1)
        addend2 = randint(1, total - addend1)
        return str(addend1) + ' + ' + str(addend2) + ' = '
    else:
        subtrahend = randint(1, total - 1)
        minuend = randint(subtrahend, total - 1)
        return str(minuend) + ' - ' + str(subtrahend) + ' = '


def create_data_3(total):
    number = randint(1, total - 1)
    result_str = str(number)
    for y in range(2):
        operation_type = randint(0, 1)
        if operation_type == 1:
            addend = randint(0, total - eval(result_str))
            result_str += ' + ' + str(addend)
        else:
            subtrahend = randint(0, eval(result_str))
            result_str += ' - ' + str(subtrahend)
    result_str += ' = '
    return result_str


def create_data(total):
    data1 = []
    data2 = []
    for x in range((total + 1) * (total + 1) * 2 * 10):
        data1.append(create_data_2(total))
    for x in range((total + 1) * (total + 1) * 3 * 10):
        data2.append(create_data_3(total))
    return list(set(data1))[:60] + list(set(data2))[:20]


input_result = int(input('请输入任意数字，计算该数字以内的加减法：'))
if input_result <= 100:
    op_to_excel(create_data(input_result), '口算题.xlsx')
else:
    print('\033[0;33;40m请输入小于等于100的数字。\033[0m')
