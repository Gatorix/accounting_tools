from datetime import date
import time
import os
import msvcrt
import xlrd
import xlwt

def format_date(book, sheet, row, col):  # 日期格式转换
    if sheet.cell(row, col).ctype == 3:
        date_value = xlrd.xldate_as_tuple(
            sheet.cell_value(row, col), book.datemode)
        date_tmp = date(*date_value[:3]).strftime('%Y-%m-%d')
    return date_tmp


def format_period(period):  # 调整期间字符串，去掉多余的0
    if period[0] == '0':
        period = period[-1]
    return period


def scientific_notation(value):  # 保留两位小数
    value_tmp = []
    try:
        for i in range(len(value)):
            value_tmp.append(round(value[i], 2))
        if 0 in value_tmp:
            print('"Data"表结算金额字段中含有"0"，删除后重试！')
            exit_with_anykey()
        return value_tmp
    except TypeError:
        print('结算金额中含有非数字. 修改后重试!')
        exit_with_anykey()


def cal_tax_free(tax_included):  # 计算不含税金额并返回
    return [round(i/1.06, 2) for i in tax_included]


def cal_tax(tax_included, tax_free):  # 计算税金并返回
    tax_tmp = [round(tax_included[i]-tax_free[i], 2)
               for i in range(len(tax_included))]
    if 0 in tax_tmp:
        print('Warning:\n', '-'*46, '\n 计算得出税金中含有"0", 凭证生成后需要手动修改！\n', '-'*46)
    return tax_tmp


def exit_with_anykey():
    print("Press any key to exit...")
    ord(msvcrt.getch())
    os._exit(1)


def check_file(input_filename):  # 检查文件是否存在
    print('Check if the file exists...')
    if os.path.exists(input_filename):
        print('Pass...')
    else:
        print("No such file!")
        print('Check if "'+input_filename +
              '" is in the same directory as the program!')
        exit_with_anykey()


def write_2d_list(sheet, t_Schema, start_row=0):
    for row in range(len(t_Schema)):
        for col in range(0, len(t_Schema[row])):
            sheet.write(row+start_row, col, t_Schema[row][col])


def get_t_Schema(sheet_r, rows):
    t_Schema = []
    for i in range(rows):
        t_Schema.append(sheet_r.row_values(i))
    return t_Schema


def write_header_line(sheet):
    header_line = [u'凭证日期', u'会计年度', u'会计期间', u'凭证字', u'凭证号', u'科目代码',
                   u'科目名称', u'币别代码', u'币别名称', u'原币金额', u'借方', u'贷方',
                   u'制单', u'审核', u'核准', u'出纳', u'经办', u'结算方式', u'结算号',
                   u'凭证摘要', u'数量', u'数量单位', u'单价', u'参考信息', u'业务日期',
                   u'往来业务编号', u'附件数', u'序号', u'系统模块', u'业务描述', u'汇率类型',
                   u'汇率', u'分录序号', u'核算项目', u'过账', u'机制凭证', u'现金流量']
    for i in range(len(header_line)):
        sheet.write(0, i, header_line[i])


def write_formula(sheet, rows, start_rows=1):
    for i in range(rows):
        start_rows = start_rows+1
        sheet.write(i+1, 9, xlwt.Formula(
            'IF(K'+str(start_rows)+'=0,L' +
            str(start_rows)+',K'+str(start_rows)+')'))


def write_entry_number(sheet, rows, entry_number=0):
    for i in range(rows):
        entry_number = entry_number+1
        sheet.write(i+1, 32, entry_number-1)


def generate_ap(ap_type, ap_code, ap_name, rows):
    new_list = ['']*rows
    list1 = [ap_type]*rows
    list2 = ['---']*rows
    try:
        for i in range(len(list1)):
            new_list[i] += list1[i]
            new_list[i] += list2[i]
            new_list[i] += ap_code[i]
            new_list[i] += list2[i]
            new_list[i] += ap_name[i]
        return new_list
    except TypeError:
        print('"Data"表中发现数据类型错误，原因可能为：\n1、公司名称字段含有纯数字；\n2、核算项目字段含有"#N/A"\n检查并修改后重试！')
        exit_with_anykey()


def generate_full_ap(rows, list1, list2):
    full_accounting_project = []
    for i in range(rows):
        full_accounting_project.append(list1[i])
        full_accounting_project.append(list2[i])
        full_accounting_project.append('')
    return full_accounting_project


def write_list(sheet, rows, cols, in_list):  # 写入list
    for i in range(rows):
        sheet.write(i+1, cols, in_list[i])


def corp_add_tax(in_list, add_str):
    for i in range(len(in_list)):
        if i % 3 == 0:
            in_list[i+2] += add_str
    return in_list


def triple_list(in_list, repeat_times=3):
    out_list = []
    for i in range(len(in_list)):
        for _ in range(repeat_times):
            out_list.append(in_list[i])
    return out_list


def reorganize_summary(VOUCHER_PERIOD, AISLE_NAME, BUSINESS_TYPE, VOUCHER_TYPE, corp_name_tax):
    voucher_summary = []
    for i in range(len(corp_name_tax)):
        voucher_summary.append('确认'+VOUCHER_PERIOD+'月'+AISLE_NAME +
                               BUSINESS_TYPE+VOUCHER_TYPE+' - '+corp_name_tax[i])
    return voucher_summary


def generate_tax_include_list(rows, in_list):
    tax_include_list = []
    for i in range(rows):
        tax_include_list.append(in_list[i])
        tax_include_list.append(0)
        tax_include_list.append(0)
    return tax_include_list


def generate_tax_free_and_tax_list(rows, list1, list2):
    tax_free_and_tax_list = []
    for i in range(rows):
        tax_free_and_tax_list.append(0)
        tax_free_and_tax_list.append(list1[i])
        tax_free_and_tax_list.append(list2[i])
    return tax_free_and_tax_list


# def formatXLS(filepath):  # 转换为xls格式
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(filepath)

#     # FileFormat = 51 is for .xlsx extension
#     # FileFormat = 56 is for .xls extension
#     wb.SaveAs(filepath[:-1], FileFormat=56)
#     wb.Close()
#     excel.Application.Quit()


print('Start...')
start_time = time.time()
INPUT_FILENAME = '待入账数据.xlsm'
check_file(INPUT_FILENAME)
workbook_r = xlrd.open_workbook(INPUT_FILENAME)  # 打开'待入账数据.xlsm'
print('Open workbook: ', workbook_r)
sheet_Configuration = workbook_r.sheet_by_name(
    'Configuration')  # 获取'Configuration'表
print('Open sheet_Configuration: ', sheet_Configuration)
sheet_Data = workbook_r.sheet_by_name('Data')  # 获取'Data'表
print('Open sheet_Data: ', sheet_Data)
sheet_t_Schema_r = workbook_r.sheet_by_name('t_Schema')  # 获取't_Schema'表
print('Open sheet_t_Schema_r: ', sheet_t_Schema_r)

nrows_Data = sheet_Data.nrows-1  # 获取有效数据行数，不含标题行
print('A total of '+str(nrows_Data)+' lines of data.')
nrows_Data_full = nrows_Data*3  # 行数*3
nrows_t_Schema = sheet_t_Schema_r.nrows
print('Reading the fixed value in the "Configuration"...')
VOUCHER_DATE = format_date(
    workbook_r, sheet_Configuration, 0, 1)  # 获取凭证日期
VOUCHER_YEAR = VOUCHER_DATE[:4]  # 获取凭证年度
VOUCHER_PERIOD = format_period(VOUCHER_DATE[5:7])  # 获取凭证期间
BOOKKEEPER = sheet_Configuration.cell_value(1, 1)  # 获取制单人
VOUCHER_TYPE = sheet_Configuration.cell_value(2, 1)  # 获取凭证类型
BUSINESS_TYPE = sheet_Configuration.cell_value(3, 1)  # 获取业务类型
AISLE_NAME = sheet_Configuration.cell_value(4, 1)  # 获取通道名称
ACCOUNTING_PROJECT_TYPE = sheet_Configuration.cell_value(5, 1)  # 获取核算项目类别
subject_code = [sheet_Configuration.cell_value(
    6, 1), sheet_Configuration.cell_value(7, 1), sheet_Configuration.cell_value(8, 1)]*nrows_Data
subject_name = [sheet_Configuration.cell_value(
    9, 1), sheet_Configuration.cell_value(10, 1), sheet_Configuration.cell_value(11, 1)]*nrows_Data
fixed_value = [VOUCHER_DATE, VOUCHER_YEAR, VOUCHER_PERIOD, '记', '1',
               '', '', 'RMB', '人民币', '', '', '', BOOKKEEPER, 'NONE',
               'NONE', 'NONE', '', '*', '', '', '0', '*', '0',
               '', VOUCHER_DATE, '', '0', '1', '', '', '公司汇率',
               '1', '', '', '0', '', ''
               ]  # 将固定字段整理为列表

print('Reading the settlement value in the "Data"...')
corp_name = sheet_Data.col_values(0, 1)  # 获取公司名称
corp_name_triple = triple_list(corp_name)
corp_name_tax = corp_add_tax(
    corp_name_triple, ' - '+sheet_Configuration.cell_value(11, 1))
voucher_summary = reorganize_summary(
    VOUCHER_PERIOD, AISLE_NAME, BUSINESS_TYPE, VOUCHER_TYPE, corp_name_tax)
tax_include_amount = scientific_notation(
    sheet_Data.col_values(1, 1))  # 获取含税金额
tax_include_amount_full = generate_tax_include_list(
    nrows_Data, tax_include_amount)
tax_free_amount = cal_tax_free(tax_include_amount)  # 计算不含税金额
tax_amount = cal_tax(tax_include_amount, tax_free_amount)  # 计算税金
tax_free_and_tax_list = generate_tax_free_and_tax_list(
    nrows_Data, tax_free_amount, tax_amount)
accounting_project_code = sheet_Data.col_values(2, 1)  # 获取核算项目代码

workbook_w = xlwt.Workbook()  # 创建文件
print('Create workbook: ', workbook_w)
sheet_Page1 = workbook_w.add_sheet(
    'Page1', cell_overwrite_ok=True)  # 新建'Page1'，可复写
print('Create sheet_Page1', sheet_Page1)
sheet_t_Schema_w = workbook_w.add_sheet('t_Schema')  # 新建't_Schema'
print('Create sheet_t_Schema_w', sheet_t_Schema_w)
t_Schema = get_t_Schema(sheet_t_Schema_r, nrows_t_Schema)  # 获取't_Schema'数据
print('Writing t_Schema...')
write_2d_list(sheet_t_Schema_w, t_Schema)  # 写入't_Schema'
print('Writing Header Line...')
write_header_line(sheet_Page1)  # 写入标题行
print('Writing Fixed Value...')
write_2d_list(sheet_Page1, [fixed_value]*nrows_Data_full, 1)  # 写入固定字段
print('Writing Formula...')
write_formula(sheet_Page1, nrows_Data_full)  # 写入原币金额
print('Writing Entry Number...')
write_entry_number(sheet_Page1, nrows_Data_full)  # 写入分录序号

accounting_project_single = generate_ap(
    ACCOUNTING_PROJECT_TYPE, accounting_project_code, corp_name, nrows_Data)
accounting_project = generate_full_ap(
    nrows_Data, accounting_project_single, accounting_project_single)

d_cols = 10 if VOUCHER_TYPE == '收入' else 11  # 根据凭证类型单元格确定借贷方向
c_cols = 11 if VOUCHER_TYPE == '收入' else 10  # 根据凭证类型单元格确定借贷方向
print('Writing Subject Code...')
write_list(sheet_Page1, nrows_Data_full, 5, subject_code)
print('Writing Subject Name...')
write_list(sheet_Page1, nrows_Data_full, 6, subject_name)
print('Writing Debit Amount...')
write_list(sheet_Page1, nrows_Data_full, d_cols, tax_include_amount_full)
print('Writing Credit Amount...')
write_list(sheet_Page1, nrows_Data_full, c_cols, tax_free_and_tax_list)
print('Writing Voucher Summary...')
write_list(sheet_Page1, nrows_Data_full, 19, voucher_summary)
print('Writing Accounting Project...')
write_list(sheet_Page1, nrows_Data_full, 33, accounting_project)
print('Saving File...')
output_filename = VOUCHER_PERIOD + '月' + \
    AISLE_NAME + BUSINESS_TYPE+VOUCHER_TYPE+'_VW.xls'
workbook_w.save(output_filename)  # 保存文件
print('File Saved in '+os.path.abspath(output_filename))
end_time = time.time()
print('Time Cost：'+str(round(end_time-start_time, 2))+'s')
exit_with_anykey()
