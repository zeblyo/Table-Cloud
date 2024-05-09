import xlrd
from datetime import datetime
from xlutils.copy import copy
import time

def is_number(s):
    try:  # 如果能运行float(s)语句，返回True（字符串s是浮点数）
        float(s)
        return True
    except ValueError:  # ValueError为Python的一种标准异常，表示"传入无效的参数"
        pass  # 如果引发了ValueError这种异常，不做任何事情（pass：不做任何事情，一般用做占位语句）
    try:
        import unicodedata  # 处理ASCii码的包
        unicodedata.numeric(s)  # 把一个表示数字的字符串转换为浮点数返回的函数
        return True
    except (TypeError, ValueError):
        pass
    return False

print("不要在打开生成表格的时候运行此程序")
print("请把此文件和所有表格放在同一文件夹内")
enemy_input = input("请输入对方公司格式的表格名称（全称）：")
user_input = input("请输入我方公司格式的表格名称（全称）：")
program_input = input("请输入自动生成表格名称（不建议于原表格相同）：")

now = datetime.now()
current_time = now.strftime("%Y-%m-%d %H:%M:%S")

with open("log.txt", "a") as file:
    file.write(current_time)


def write_log(str):
    with open("log.txt", "a") as file:
        file.write(str)

# 打开Excel文件
workbook = xlrd.open_workbook(f'./{enemy_input}')
workbook1 = xlrd.open_workbook(f'./{user_input}')

# 获取第一个工作表
worksheet = workbook.sheet_by_index(0)
worksheet1 = workbook1.sheet_by_index(0)

# 读取整列的值
column_values = worksheet.col_values(7, 3)
column_values1 = worksheet1.col_values(2, 2)

new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象

# 写入表格信息
write_save = new_workbook.get_sheet(0)

for n, i in enumerate(column_values):
    if i in column_values1:
        value = worksheet1.cell_value(column_values1.index(i)+2, 5)
        value1 = worksheet1.cell_value(column_values1.index(i)+2, 6)
        write_save.write(n+3, 12, f"{value}*{value1}")

        num = worksheet.cell_value(n+3, 4)
        num1 = worksheet1.cell_value(column_values1.index(i)+2, 3)
        if num != num1 and is_number(num) and is_number(num1):
            print(f"对方公司表格中第{n+4}行{num}的收货量有错，对应我们的表格中第{column_values1.index(i)+3}行{num1}的收货量")
            write_log(f"\n对方公司表格中第{n+4}行{num}的收货量有错，对应我们的表格中第{column_values1.index(i)+3}行{num1}的收货量")
        
        sale = worksheet.cell_value(n+3, 5)
        sale1 = worksheet1.cell_value(column_values1.index(i)+2, 7)
        if is_number(sale) and is_number(sale1):
            diff = float(sale) - float(sale1)
            if abs(diff) > 0.1:
                print(f"对方公司表格中第{n+4}行{sale}的总价有错，对应我们的表格中第{column_values1.index(i)+3}行{sale1}的总价")
                write_log(f"\n对方公司表格中第{n+4}行{sale}的总价有错，对应我们的表格中第{column_values1.index(i)+3}行{sale1}的总价")
    else:
        print(f"元素 {i} 不在列表 2 中")
        write_log(f"\n元素 {i} 不在列表 2 中")

new_workbook.save(f"{program_input}.xls")  # 保存工作簿

time.sleep(99999999)