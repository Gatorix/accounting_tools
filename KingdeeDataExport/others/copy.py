import os
import xlrd
import xlwt


def open_excel(input_filename, index=0):  # 打开要解析的Excel文件
    try:
        data = xlrd.open_workbook(input_filename)
        table = data.sheets()[index]
        return table
    except FileNotFoundError:
        print("No such file: "+input_filename)
        print("Press any key to exit...")
        ord(msvcrt.getch())
        os._exit(1)


def listdir(path, list_name):
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if os.path.isdir(file_path):
            listdir(file_path, list_name)
        elif os.path.splitext(file_path)[1] == '.xls':
            list_name.append(file_path)


def write_2d_value(sheet, result):
    for c in range(len(result)):
        for r in range(0, len(result[c])):
            sheet.write(r, c, result[c][r])


file_list = []
listdir("C:/Users/ZCY-CW/Desktop/瑞华年审资料-第三批/进销存数据0220/客户", file_list)


book = xlwt.Workbook()
for i in range(12):
    sheet = open_excel(file_list[i])
    col_value1 = []
    for x in range(2):
        col_value1.append(sheet.col_values(x))
    for y in range(3):
        col_value1.append(sheet.col_values(6+y))
# print(col_value1)

    sh = book.add_sheet(str(i+1)+'月')
    write_2d_value(sh, col_value1)
book.save('客户汇总.xls')
