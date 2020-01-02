from operatingxls import ReadXLS, WriteXLS
import xlwt

voucher_date = ReadXLS.get_cols_value(0, 0)  # 取日期

voucher_date.remove('日期')
# # print(accounting_project_code)


def write_value(voucher_date):
    date_style = xlwt.XFStyle()
    date_style.num_format_str = 'yyyy/mm/dd'
    for i in range(0, len(voucher_date)):
        WriteXLS.sheet1.write(i+1, 0, voucher_date[i], date_style)
    WriteXLS.workbook.save('test.xls')


# write_value(voucher_date)

# def check_input_file_cols(input_filename):
#     print("Check if the number of input file cols is correct...")
#     if input_filename == '待入账数据-结算.xls':
#         if ReadXLS.get_cols(input_filename) == 5:
#             print('Done...')
#         else:
#             print('File data is incorrect!')
#     elif input_filename == '待入账数据-银行.xls':
#         if ReadXLS.get_cols(input_filename) == 6:
#             print('Done...')
#         else:
#             print('File data is incorrect!')


# # check_input_file_cols()
# advance_receipt = ReadXLS.get_cols_value(1)  # 取含税金额
# # advance_receipt.remove('含税金额')
# print(advance_receipt)
