import generater.funcs as funcs
from generater.operatingxls import WriteXLS, ReadXLS
import time


def income_generate():
    funcs.check_file(ReadXLS.input_filename)
    WriteXLS.save_file()
    voucher_date = funcs.get_voucher_date()
    business_date = voucher_date  # 业务日期
    voucher_bookkeeper = funcs.get_voucher_bookkeeper()  # 判断输入
    business_type, channel_name = funcs.get_business_type()
    print('\nStart generating voucher template...')
    start_time = time.time()
    voucher_year = funcs.get_voucher_year(voucher_date)  # 会计年度
    voucher_period = funcs.get_voucher_period(voucher_date)  # 会计期间
    print('Reading input file...')
    corp_name = ReadXLS.get_cols_value(0)  # 取公司名称
    corp_name.remove('公司名称')
    advance_receipt = ReadXLS.get_cols_value(1)  # 取含税金额
    advance_receipt.remove('含税金额')
    income_amount = ReadXLS.get_cols_value(2)  # 取结算金额
    income_amount.remove('结算金额')
    tax_amount = ReadXLS.get_cols_value(3)  # 取税金金额
    tax_amount.remove('税金金额')
    accounting_project_code = ReadXLS.get_cols_value(4)  # 取核算项目编码
    accounting_project_code.remove('核算项目代码')
    WriteXLS.write_fixed_value(
        voucher_date, voucher_year, voucher_period, voucher_bookkeeper, business_date)  # 写入固定值
    WriteXLS.write_voucher_summary_income(
        corp_name, voucher_period, business_type, channel_name)  # 写入摘要
    WriteXLS.write_debit_credit_amount_income(
        advance_receipt, income_amount, tax_amount)  # 写入借贷金额
    subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount, subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount = funcs.get_income_subject(
        business_type)  # 获取业务类型对应的科目代码及科目名称
    WriteXLS.write_subject(subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount,
                           subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount)  # 写入科目代码及科目名称
    WriteXLS.write_accounting_project_income(
        accounting_project_code, corp_name)  # 写入核算项目
    WriteXLS.write_t_Schema()  # 写入参数表
    end_time = time.time()
    print('Time cost', round(end_time-start_time, 2), 's...')


def cost_generate():
    funcs.check_file(ReadXLS.input_filename)
    WriteXLS.save_file()
    voucher_date = funcs.get_voucher_date()
    business_date = voucher_date  # 业务日期
    voucher_bookkeeper = funcs.get_voucher_bookkeeper()  # 判断输入
    business_type, channel_name = funcs.get_business_type()
    print('\nStart generating voucher template...')
    start_time = time.time()
    voucher_year = funcs.get_voucher_year(voucher_date)  # 会计年度
    voucher_period = funcs.get_voucher_period(voucher_date)  # 会计期间
    print('Reading input file...')
    corp_name = ReadXLS.get_cols_value(0)  # 取公司名称
    corp_name.remove('公司名称')
    advance_receipt = ReadXLS.get_cols_value(1)  # 取预收金额
    advance_receipt.remove('含税金额')
    income_amount = ReadXLS.get_cols_value(2)  # 取收入金额
    income_amount.remove('结算金额')
    tax_amount = ReadXLS.get_cols_value(3)  # 取税金金额
    tax_amount.remove('税金金额')
    accounting_project_code = ReadXLS.get_cols_value(4)  # 取核算项目编码
    accounting_project_code.remove('核算项目代码')
    WriteXLS.write_fixed_value(
        voucher_date, voucher_year, voucher_period, voucher_bookkeeper, business_date)  # 写入固定值
    WriteXLS.write_voucher_summary_cost(
        corp_name, voucher_period, business_type, channel_name)  # 写入摘要
    WriteXLS.write_debit_credit_amount_cost(
        advance_receipt, income_amount, tax_amount)  # 写入借贷金额
    subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount, subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount = funcs.get_cost_subject(
        business_type)  # 获取业务类型对应的科目代码及科目名称
    WriteXLS.write_subject(subject_code_advance_receipt, subject_name_advance_receipt, subject_code_income_amount,
                           subject_name_income_amount, subject_code_tax_amount, subject_name_tax_amount)  # 写入科目代码及科目名称
    WriteXLS.write_accounting_project_cost(
        accounting_project_code, corp_name)  # 写入核算项目
    WriteXLS.write_t_Schema()  # 写入参数表
    end_time = time.time()
    print('Time cost', round(end_time-start_time, 2), 's...')


def bank_generate():
    # voucher_bookkeeper = funcs.get_voucher_bookkeeper()  # 判断输入
    print('bank_income_generate')
