import xlrd
import xlwt
import msvcrt
import os


class ReadXLS():

    input_filename = "./待入账数据.xls"

    def open_excel(input_filename, index=0):  # 打开要解析的Excel文件
        try:
            data = xlrd.open_workbook(input_filename)
            return data
        except FileNotFoundError:
            print("No such file: "+input_filename)
            print("Press any key to exit...")
            ord(msvcrt.getch())
            os._exit(1)

    def get_cols(input_filename, by_index=0):  # 获取列数
        data = ReadXLS.open_excel(input_filename)  # 打开excel文件
        tab = data.sheets()[by_index]  # 选择excel里面的Sheet
        ncols = tab.ncols  # 列数
        return ncols

    def get_cols_value(ncols, by_index=0):  # 获取列值
        data = ReadXLS.open_excel(ReadXLS.input_filename)  # 打开excel文件
        sheetx = data.sheets()[by_index]  # 选择excel里面的Sheet
        return sheetx.col_values(ncols)

    def get_rows(input_filename, by_index=0):  # 获取行数
        try:
            data = ReadXLS.open_excel(input_filename)  # 打开excel文件
            tab = data.sheets()[by_index]  # 选择excel里面的Sheet
            nrows = tab.nrows  # 行数
            return nrows
        except AttributeError:
            print("No such file: "+input_filename)
            print("Press any key to exit...")
            ord(msvcrt.getch())
            os._exit(1)

    # def get_rows_value


class WriteXLS():
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet('Page1', cell_overwrite_ok=True)
    sheet2 = workbook.add_sheet('t_Schema', cell_overwrite_ok=True)
    input_filename = ReadXLS.input_filename
    output_filename = "./output.xls"
    settle_rows = ReadXLS.get_rows(input_filename)-1

    def save_file():
        print('Check if the output file is writable...')
        try:
            WriteXLS.workbook.save(WriteXLS.output_filename)
            print('Done')
        except PermissionError:
            print('Permission Denied')
            print("Press any key to exit...")
            ord(msvcrt.getch())
            os._exit(1)

    def write_fixed_value(voucher_date, voucher_year, voucher_period,
                          voucher_bookkeeper, business_date, entry_number=0, start_cows=1):  # 写入固定值

        print('Writing header line...')
        field_list = [u'凭证日期', u'会计年度', u'会计期间', u'凭证字', u'凭证号', u'科目代码',
                      u'科目名称', u'币别代码', u'币别名称', u'原币金额', u'借方', u'贷方',
                      u'制单', u'审核', u'核准', u'出纳', u'经办', u'结算方式', u'结算号',
                      u'凭证摘要', u'数量', u'数量单位', u'单价', u'参考信息', u'业务日期',
                      u'往来业务编号', u'附件数', u'序号', u'系统模块', u'业务描述', u'汇率类型',
                      u'汇率', u'分录序号', u'核算项目', u'过账', u'机制凭证', u'现金流量']
        for i in range(0, len(field_list)):
            WriteXLS.sheet1.write(0, i, field_list[i])

        print('Writing fixed field...')
        voucher_symbol = "记"  # 凭证字
        voucher_number = 1  # 凭证号
        currency_code = "RMB"  # 币别代码
        currency_name = "人民币"  # 币别名称
        voucher_reviewer = "NONE"  # 审核
        voucher_approver = "NONE"  # 核准
        voucher_cashier = "NONE"  # 出纳
        voucher_attn = ""  # 经办
        settlement_method = "*"  # 结算方式
        settlement_number = ""  # 结算号
        voucher_quantity = 0  # 数量
        quantity_unit = "*"  # 数量单位
        unit_price = 0  # 单价
        reference_information = ""  # 参考信息
        business_dealings = ""  # 往来业务编号
        attachments_number = 0  # 附件数
        serial_number = 1  # 序号
        system_module = ""  # 系统模块
        business_description = ""  # 业务描述
        rate_type = "公司汇率"  # 公司汇率
        voucher_rate = 1  # 汇率
        voucher_posting = 0
        redential_mechanism = ""  # 凭证机制
        cash_flow = ""  # 现金流量
        for i in range(0, 3*WriteXLS.settle_rows):
            WriteXLS.sheet1.write(i+1, 0, voucher_date)
            WriteXLS.sheet1.write(i+1, 1, voucher_year)
            WriteXLS.sheet1.write(i+1, 2, voucher_period)
            WriteXLS.sheet1.write(i+1, 3, voucher_symbol)
            WriteXLS.sheet1.write(i+1, 4, voucher_number)
            WriteXLS.sheet1.write(i+1, 7, currency_code)
            WriteXLS.sheet1.write(i+1, 8, currency_name)
            WriteXLS.sheet1.write(i+1, 12, voucher_bookkeeper)
            WriteXLS.sheet1.write(i+1, 13, voucher_reviewer)
            WriteXLS.sheet1.write(i+1, 14, voucher_approver)
            WriteXLS.sheet1.write(i+1, 15, voucher_cashier)
            WriteXLS.sheet1.write(i+1, 16, voucher_attn)
            WriteXLS.sheet1.write(i+1, 17, settlement_method)
            WriteXLS.sheet1.write(i+1, 18, settlement_number)
            WriteXLS.sheet1.write(i+1, 20, voucher_quantity)
            WriteXLS.sheet1.write(i+1, 21, quantity_unit)
            WriteXLS.sheet1.write(i+1, 22, unit_price)
            WriteXLS.sheet1.write(i+1, 23, reference_information)
            WriteXLS.sheet1.write(i+1, 24, business_date)
            WriteXLS.sheet1.write(i+1, 25, business_dealings)
            WriteXLS.sheet1.write(i+1, 26, attachments_number)
            WriteXLS.sheet1.write(i+1, 27, serial_number)
            WriteXLS.sheet1.write(i+1, 28, system_module)
            WriteXLS.sheet1.write(i+1, 29, business_description)
            WriteXLS.sheet1.write(i+1, 30, rate_type)
            WriteXLS.sheet1.write(i+1, 31, voucher_rate)
            WriteXLS.sheet1.write(i+1, 34, voucher_posting)
            WriteXLS.sheet1.write(i+1, 35, redential_mechanism)
            WriteXLS.sheet1.write(i+1, 36, cash_flow)

        print('Writing entry number...')
        for i in range(0, 3*WriteXLS.settle_rows):
            entry_number = entry_number+1
            WriteXLS.sheet1.write(i+1, 32, entry_number-1)

        print('Writing the original currency amount...')
        for i in range(0, 3*WriteXLS.settle_rows):
            start_cows = start_cows+1
            WriteXLS.sheet1.write(i+1, 9, xlwt.Formula(
                'IF(K'+str(start_cows)+'=0,L' +
                str(start_cows)+',K'+str(start_cows)+')'))

        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_voucher_summary_income(corp_name, voucher_period,
                                     business_type, channel_name):  # 写入收入类摘要
        print('Writing summary...')
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+1, 19, '结算'+voucher_period+'月收入-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i])
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+len(corp_name)+1, 19, '结算'+voucher_period +
                                  '月收入-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i])
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+2*len(corp_name)+1, 19, '结算'+voucher_period
                                  + '月收入-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i]+'-已结算未开发票销项税额')
        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_voucher_summary_cost(corp_name, voucher_period,
                                   business_type, channel_name):  # 写入成本类摘要
        print('Writing summary...')
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+1, 19, '结算'+voucher_period+'月成本-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i])
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+len(corp_name)+1, 19, '结算'+voucher_period +
                                  '月成本-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i])
        for i in range(0, len(corp_name)):
            WriteXLS.sheet1.write(i+2*len(corp_name)+1, 19, '结算'+voucher_period
                                  + '月成本-' +
                                  channel_name+business_type +
                                  '业务-'+corp_name[i]+'-已结算未开发票进项税额')
        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_debit_credit_amount_income(advance_receipt,
                                         income_amount, tax_amount):  # 写入借贷金额
        print('Writing debit/credit amount...')
        for i in range(0, len(advance_receipt)):
            WriteXLS.sheet1.write(i+1, 10, advance_receipt[i])
        for i in range(0, len(income_amount)):
            WriteXLS.sheet1.write(
                i+len(advance_receipt)+1, 11, income_amount[i])
        for i in range(0, len(tax_amount)):
            WriteXLS.sheet1.write(
                i+2*len(advance_receipt)+1, 11, tax_amount[i])
        for i in range(0, len(advance_receipt)):
            WriteXLS.sheet1.write(i+1, 11, 0)
        for i in range(0, len(income_amount)):
            WriteXLS.sheet1.write(i+len(advance_receipt)+1, 10, 0)
        for i in range(0, len(tax_amount)):
            WriteXLS.sheet1.write(i+2*len(advance_receipt)+1, 10, 0)
        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_debit_credit_amount_cost(advance_receipt,
                                       income_amount, tax_amount):  # 写入借贷金额
        print('Writing debit/credit amount...')
        for i in range(0, len(advance_receipt)):
            WriteXLS.sheet1.write(i+1, 11, advance_receipt[i])
        for i in range(0, len(income_amount)):
            WriteXLS.sheet1.write(
                i+len(advance_receipt)+1, 10, income_amount[i])
        for i in range(0, len(tax_amount)):
            WriteXLS.sheet1.write(
                i+2*len(advance_receipt)+1, 10, tax_amount[i])
        for i in range(0, len(advance_receipt)):
            WriteXLS.sheet1.write(i+1, 10, 0)
        for i in range(0, len(income_amount)):
            WriteXLS.sheet1.write(i+len(advance_receipt)+1, 11, 0)
        for i in range(0, len(tax_amount)):
            WriteXLS.sheet1.write(i+2*len(advance_receipt)+1, 11, 0)
        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_subject(subject_code_advance_receipt,
                      subject_name_advance_receipt,
                      subject_code_income_amount,
                      subject_name_income_amount,
                      subject_code_tax_amount,
                      subject_name_tax_amount):  # 写入科目名称及科目代码
        print('Writing subject name and subject code...')
        settle_rows = WriteXLS.settle_rows
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(i+1, 5, subject_code_advance_receipt)
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(i+1, 6, subject_name_advance_receipt)
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(
                i+settle_rows+1, 5, subject_code_income_amount)
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(
                i+settle_rows+1, 6, subject_name_income_amount)
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(
                i+2*settle_rows+1, 5, subject_code_tax_amount)
        for i in range(0, settle_rows):
            WriteXLS.sheet1.write(
                i+2*settle_rows+1, 6, subject_name_tax_amount)
        WriteXLS.workbook.save(WriteXLS.output_filename)

    def write_accounting_project_income(accounting_project_code, corp_name):  # 写入核算项目
        print('Writing accounting project...')
        settle_rows = WriteXLS.settle_rows
        try:
            for i in range(0, settle_rows):
                WriteXLS.sheet1.write(i+1, 33, '客户---' +
                                      accounting_project_code[i]+'---' +
                                      corp_name[i])
            for i in range(0, settle_rows):
                WriteXLS.sheet1.write(i+settle_rows+1, 33, '客户---' +
                                      accounting_project_code[i]+'---' +
                                      corp_name[i])
            WriteXLS.workbook.save(WriteXLS.output_filename)
        except Exception:
            print('Value Error!')
            print("Press any key to exit...")
            ord(msvcrt.getch())
            os._exit(1)

    def write_accounting_project_cost(accounting_project_code, corp_name):  # 写入核算项目
        print('Writing accounting project...')
        settle_rows = WriteXLS.settle_rows
        try:
            for i in range(0, settle_rows):
                WriteXLS.sheet1.write(i+1, 33, '供应商---' +
                                      accounting_project_code[i]+'---' +
                                      corp_name[i])
            for i in range(0, settle_rows):
                WriteXLS.sheet1.write(i+settle_rows+1, 33, '供应商---' +
                                      accounting_project_code[i]+'---' +
                                      corp_name[i])
            WriteXLS.workbook.save(WriteXLS.output_filename)
        except Exception:
            print('Value Error!')
            print("Press any key to exit...")
            ord(msvcrt.getch())
            os._exit(1)

    def write_t_Schema():  # 写入金蝶模板参数表
        print('Writing template parameter table...')
        value = [[u'FType', u'FKey', u'FFieldName', u'FCaption', u'FValueType',
                  u'FNeedSave', u'FColIndex', u'FSrcTableName', u'FSrcFieldName', u'FExpFieldName',
                  u'FImpFieldName', u'FDefaultVal', u'FSearch', u'FItemPageName',
                  u'FTrueType', u'FPrecision', u'FSearchName', u'FIsShownList', u'FViewMask', u'FPage'],
                 [u'ClassInfo', u'ClassType', u'', u'VoucherData', u'', u'',
                  u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u''],
                 [u'ClassInfo', u'ClassTypeID', u'', u' 123', u'', u'', u'',
                  u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u''],
                 [u'PageInfo', u'Page1', u'', u'Page1', u'', u'', u'', u'',
                  u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u'', u''],
                 [u'FieldInfo', u'FDate', u'FDate', u'凭证日期', u' DateTime', u'', u'1', u'',
                  u'', u'FDate', u'FDate', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FYear', u'FYear', u'会计年度',
                  u' Decimal(28,10)', u'', u'2', u'', u'', u'FYear', u'FYear', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FPeriod', u'FPeriod', u'会计期间',
                  u' Decimal(28,10)', u'', u'3', u'', u'', u'FPeriod', u'FPeriod', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FGroupID', u'FGroupID', u'凭证字',
                  u' Varchar(80)', u'', u'4', u'', u'', u'FGroupID', u'FGroupID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FNumber', u'FNumber', u'凭证号',
                  u' Decimal(28,10)', u'', u'5', u'', u'', u'FNumber', u'FNumber', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FAccountNum', u'FAccountNum', u'科目代码',
                  u' Varchar(40)', u'', u'7', u'', u'', u'FAccountNum', u'FAccountNum', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FAccountName', u'FAccountName', u'科目名称',
                  u' Varchar(255)', u'', u'8', u'', u'', u'FAccountName', u'FAccountName', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCurrencyNum', u'FCurrencyNum', u'币别代码',
                  u' Varchar(10)', u'', u'9', u'', u'', u'FCurrencyNum', u'FCurrencyNum', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCurrencyName', u'FCurrencyName', u'币别名称',
                  u' Varchar(40)', u'', u'10', u'', u'', u'FCurrencyName', u'FCurrencyName', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FAmountFor', u'FAmountFor', u'原币金额',
                  u' Decimal(28,10)', u'', u'11', u'', u'', u'FAmountFor', u'FAmountFor', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FDebit', u'FDebit', u'借方',
                  u' Decimal(28,10)', u'', u'12', u'', u'', u'FDebit', u'FDebit', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCredit', u'FCredit', u'贷方',
                  u' Decimal(28,10)', u'', u'13', u'', u'', u'FCredit', u'FCredit', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FPreparerID', u'FPreparerID', u'制单',
                  u' Varchar(255)', u'', u'14', u'', u'', u'FPreparerID', u'FPreparerID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCheckerID', u'FCheckerID', u'审核',
                  u' Varchar(255)', u'', u'15', u'', u'', u'FCheckerID', u'FCheckerID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FApproveID', u'FApproveID', u'核准',
                  u' Varchar(255)', u'', u'17', u'', u'', u'FApproveID', u'FApproveID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCashierID', u'FCashierID', u'出纳',
                  u' Varchar(255)', u'', u'18', u'', u'', u'FCashierID', u'FCashierID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FHandler', u'FHandler', u'经办',
                  u' Varchar(50)', u'', u'19', u'', u'', u'FHandler', u'FHandler', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FSettleTypeID', u'FSettleTypeID', u'结算方式',
                  u' Varchar(80)', u'', u'20', u'', u'', u'FSettleTypeID', u'FSettleTypeID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FSettleNo', u'FSettleNo', u'结算号',
                  u' Varchar(255)', u'', u'21', u'', u'', u'FSettleNo', u'FSettleNo', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FExplanation', u'FExplanation', u'凭证摘要',
                  u' Varchar(255)', u'', u'22', u'', u'', u'FExplanation', u'FExplanation', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FQuantity', u'FQuantity', u'数量',
                  u' Decimal(28,10)', u'', u'23', u'', u'', u'FQuantity', u'FQuantity', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FMeasureUnitID', u'FMeasureUnitID', u'数量单位',
                  u' Varchar(255)', u'', u'24', u'', u'', u'FMeasureUnitID', u'FMeasureUnitID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FUnitPrice', u'FUnitPrice', u'单价',
                  u' Decimal(28,10)', u'', u'25', u'', u'', u'FUnitPrice', u'FUnitPrice', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FReference', u'FReference', u'参考信息',
                  u' Varchar(255)', u'', u'26', u'', u'', u'FReference', u'FReference', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FTransDate', u'FTransDate', u'业务日期', u' DateTime', u'', u'27',
                  u'', u'', u'FTransDate', u'FTransDate', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FTransNo', u'FTransNo', u'往来业务编号',
                  u' Varchar(255)', u'', u'28', u'', u'', u'FTransNo', u'FTransNo', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FAttachments', u'FAttachments', u'附件数',
                  u' Decimal(28,10)', u'', u'29', u'', u'', u'FAttachments', u'FAttachments', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FSerialNum', u'FSerialNum', u'序号',
                  u' Decimal(28,10)', u'', u'30', u'', u'', u'FSerialNum', u'FSerialNum', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FObjectName', u'FObjectName', u'系统模块',
                  u' Varchar(100)', u'', u'31', u'', u'', u'FObjectName', u'FObjectName', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FParameter', u'FParameter', u'业务描述',
                  u' Varchar(100)', u'', u'32', u'', u'', u'FParameter', u'FParameter', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FExchangeRateType', u'FExchangeRateType', u'汇率类型',
                  u' Varchar(80)', u'', u'33', u'', u'', u'FExchangeRateType', u'FExchangeRateType', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FExchangeRate', u'FExchangeRate', u'汇率',
                  u' Decimal(28,10)', u'', u'34', u'', u'', u'FExchangeRate', u'FExchangeRate', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FEntryID', u'FEntryID', u'分录序号',
                  u' Decimal(28,10)', u'', u'35', u'', u'', u'FEntryID', u'FEntryID', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FItem', u'FItem', u'核算项目', u'Text', u'', u'36', u'',
                  u'', u'FItem', u'FItem', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FPosted', u'FPosted', u'过账',
                  u' Decimal(28,10)', u'', u'37', u'', u'', u'FPosted', u'FPosted', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FInternalInd', u'FInternalInd', u'机制凭证',
                  u' Varchar(10)', u'', u'38', u'', u'', u'FInternalInd', u'FInternalInd', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1'],
                 [u'FieldInfo', u'FCashFlow', u'FCashFlow', u'现金流量', u'Text', u'', u'39', u'',
                  u'', u'FCashFlow', u'FCashFlow', u'', u'0', u'', u'', u'0', u'0', u'0', u'0', u'1']]

        valuerow = len(value)

        for row in range(valuerow):
            for col in range(0, len(value[row])):
                WriteXLS.sheet2.write(row, col, value[row][col])

        WriteXLS.workbook.save(WriteXLS.output_filename)
