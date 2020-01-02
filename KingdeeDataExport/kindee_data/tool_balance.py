# coding=utf-8
from . import core_sql_statement as sqls
from . import core_operatingXLS as optx
from . import core_other_funcs as funcs


def get_balance(db_name_needed, company_name_needed, msg, fyear, fperiod, fnumber_begin, fnumber_end, flevel):
    funcs.change_dir(funcs.mkdir(fyear+'年'+fperiod+'月_科目余额表'))
    for i in range(len(db_name_needed)):
        print('正在处理第['+str(i+1)+']个账套：'+company_name_needed[i]+'...')
        book = optx.create_book()  # 创建excel
        balance_sheet = book.add_sheet('科目余额表')
        result = msg.ExecQuery(
            sqls.sql_balance(db_name_needed[i], fyear, fperiod, fnumber_begin, fnumber_end, flevel))
        optx.write_header_line_trail_balance(balance_sheet)
        optx.write_db_result_trail_balance(balance_sheet, result)
        book.save('科目余额表_'+fyear+'年'+fperiod+'月_'+company_name_needed[i]+'.xls')
