import time
import os
import msvcrt
import re
import generater.auto as auto


def get_voucher_date():
    while True:
        voucher_date = input("\n输入凭证日期(YYYY-mm-dd)：")
        if logic_voucher_date(voucher_date):
            break
    return voucher_date


def get_voucher_bookkeeper():
    while True:
        voucher_bookkeeper = input("\n输入制单人：")
        if logic_voucher_bookkeeper(voucher_bookkeeper):
            break
    return voucher_bookkeeper


def get_business_type():
    while True:
        business_type = input("\n输入该凭证的业务类型(流量、语音、物联网)：")
        channel_name = ''
        if business_type == '物联网':
            channel_name = input('\n输入通道名称：')
        if logic_business_type(business_type):
            break
    return business_type, channel_name


def get_voucher_type():
    while True:
        voucher_type = input("\n输入凭证类型(收入、成本)：")
        if logic_voucher_type(voucher_type):
            break
    return voucher_type


def logic_voucher_date(voucher_date):  # 获取用户录入的日期，判断是否合法，合法返回True，否则提示“输入错误”
    try:
        time.strptime(voucher_date, "%Y-%m-%d")
        return True
    except ValueError:
        print("\ninput error!")


def logic_voucher_bookkeeper(bookkeeper):  # 用户输入制单人,判断输入的制单人是否存在于列表中
    LIST_BOOKKEEPER = ["Administrator", "曹圣", "黎光波", "林嘉铭", "王静", "张哲"]
    if bookkeeper in LIST_BOOKKEEPER:
        return True
    else:
        print("\ninput error!")


def logic_business_type(business_type):  # 用户输入业务类型，判断是否合规
    LIST_BUSINESS_TYPE = ["流量", "语音", "物联网"]
    if business_type in LIST_BUSINESS_TYPE:
        return True
    else:
        print("\ninput error!")


def logic_voucher_type(voucher_type):
    LIST_VOUCHER_TYPE = ['收入', '成本', '银行']
    if voucher_type in LIST_VOUCHER_TYPE:
        return True
    else:
        print("\ninput error!")


def decide_generate_method(voucher_type):
    if voucher_type == '收入':
        auto.income_generate()
    elif voucher_type == '成本':
        auto.cost_generate()
    elif voucher_type == '银行':
        auto.bank_generate()


def get_voucher_period(voucher_period):  # 从用户输入的日期中截取月份
    rg = re.compile(r'-(.*?)-', re.IGNORECASE | re.DOTALL)
    period = rg.search(voucher_period)
    if period:
        voucher_period = period.group(1)
    return voucher_period


def get_voucher_year(voucher_year):  # 从用户输入的日期中截取年份
    rg = re.compile(r'(.*?)-', re.IGNORECASE | re.DOTALL)
    year = rg.search(voucher_year)
    if year:
        voucher_year = year.group(1)
    return voucher_year


def get_income_subject(business_type):
    if business_type == '流量':
        subject_code_advance_receipt = "2203.01.02.01"
        subject_name_advance_receipt = "后向流量经营业务"
        subject_code_income_amount = "6001.02.01"
        subject_name_income_amount = "后向流量经营业务"
        subject_code_tax_amount = "2221.01.05"
        subject_name_tax_amount = "暂估销项税额"
    elif business_type == '语音':
        subject_code_advance_receipt = "2203.01.02.02"
        subject_name_advance_receipt = "语音业务"
        subject_code_income_amount = "6001.02.02"
        subject_name_income_amount = "语音业务"
        subject_code_tax_amount = "2221.01.05"
        subject_name_tax_amount = "暂估销项税额"
    elif business_type == '物联网':
        subject_code_advance_receipt = "2203.01.02.04"
        subject_name_advance_receipt = "智能物联云服务业务"
        subject_code_income_amount = "6001.02.03"
        subject_name_income_amount = "智能物联云服务业务"
        subject_code_tax_amount = "2221.01.05"
        subject_name_tax_amount = "暂估销项税额"
    return subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount, subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount


def get_cost_subject(business_type):
    if business_type == '流量':
        subject_code_advance_receipt = "1123.01.02.01"
        subject_name_advance_receipt = "后向流量经营业务"
        subject_code_income_amount = "6401.02.04.01"
        subject_name_income_amount = "后向流量经营业务"
        subject_code_tax_amount = "2221.01.07"
        subject_name_tax_amount = "暂估进项税"
    elif business_type == '语音':
        subject_code_advance_receipt = "1123.01.02.02"
        subject_name_advance_receipt = "语音业务"
        subject_code_income_amount = "6401.02.05.01"
        subject_name_income_amount = "采购成本"
        subject_code_tax_amount = "2221.01.07"
        subject_name_tax_amount = "暂估进项税"
    elif business_type == '物联网':
        subject_code_advance_receipt = "1123.01.02.03"
        subject_name_advance_receipt = "智能物联云服务业务"
        subject_code_income_amount = "6401.02.03.01"
        subject_name_income_amount = "采购成本"
        subject_code_tax_amount = "2221.01.07"
        subject_name_tax_amount = "暂估进项税"
    return subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount, subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount


def check_file(input_filename):  # 检查文件是否存在
    print('\nCheck if the input file exists...')
    if os.path.exists(input_filename):
        print('Done')
        pass
    else:
        print("file dosen't exist!")
        exit_with_anykey()


def exit_with_anykey():
    print("press any key to exit")
    ord(msvcrt.getch())
    os._exit(1)
