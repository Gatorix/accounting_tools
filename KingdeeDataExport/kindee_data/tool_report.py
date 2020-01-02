# coding=utf-8
from . import core_sql_statement as sqls
from . import core_operatingXLS as optx
from . import core_other_funcs as funcs


def get_report(db_name_needed, company_name_needed, msg, fyear, fperiod):
    funcs.change_dir(funcs.mkdir('财务报表_'+fyear+'年'+fperiod+'月'))
    for i in range(len(db_name_needed)):
        print('正在处理第['+str(i+1)+']个账套：'+company_name_needed[i]+'...')
        book = optx.create_book()  # 创建excel
        sheet1 = book.add_sheet('资产负债表', True)
        sheet2 = book.add_sheet('利润表', True)
        optx.write_balance_sheet_fixed(
            sheet1, company_name_needed[i], fyear, fperiod)
        optx.write_profit_statement_fixed(
            sheet2, company_name_needed[i], fyear, fperiod)
        # 期初货币资金
        cash_at_bank_and_on_hand_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "(FNumber='1001' OR FNumber='1002' OR FNumber='1012')"))
        optx.format_num_style(
            sheet1, 4, 2, cash_at_bank_and_on_hand_begin[0][0])
        # 期末货币资金
        cash_at_bank_and_on_hand_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "(FNumber='1001' OR FNumber='1002' OR FNumber='1012')"))
        optx.format_num_style(
            sheet1, 4, 1, cash_at_bank_and_on_hand_end[0][0])

        # 期初交易性金融资产
        financial_assets_held_for_trading_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1101'"))
        optx.format_num_style(
            sheet1, 5, 2, financial_assets_held_for_trading_begin[0][0])
        # 期末交易性金融资产
        financial_assets_held_for_trading_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1101'"))
        optx.format_num_style(
            sheet1, 5, 1, financial_assets_held_for_trading_end[0][0])

        # 期初衍生金融资产
        optx.format_num_style(
            sheet1, 6, 2, 0)
        # 期末衍生金融资产
        optx.format_num_style(
            sheet1, 6, 1, 0)

        # 期初应收票据及应收账款
        notes_receivable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1121'"))
        accounts_receivable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1122'"))
        doubtful_debts_provision_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1231.01'", flevel='2'))
        optx.format_num_style(
            sheet1, 7, 2, notes_receivable_begin[0][0]+accounts_receivable_begin[0][0]+doubtful_debts_provision_begin[0][0])
        # 期末应收票据及应收账款
        notes_receivable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1121'"))
        accounts_receivable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1122'"))
        doubtful_debts_provision_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1231.01'", flevel='2'))
        optx.format_num_style(
            sheet1, 7, 1, notes_receivable_end[0][0]+accounts_receivable_end[0][0]+doubtful_debts_provision_end[0][0])

        # 期初预付账款
        advances_to_suppliers_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1123'"))
        optx.format_num_style(
            sheet1, 8, 2, advances_to_suppliers_begin[0][0])
        # 期末预付账款
        advances_to_suppliers_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1123'"))
        optx.format_num_style(
            sheet1, 8, 1, advances_to_suppliers_end[0][0])

        # 期初其他应收款
        other_receivables_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1221'"))
        doubtful_debts_provision_o_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1231.02'", flevel='2'))
        optx.format_num_style(
            sheet1, 9, 2, other_receivables_begin[0][0]+doubtful_debts_provision_o_begin[0][0])
        # 期末其他应收款
        other_receivables_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1221'"))
        doubtful_debts_provision_o_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1231.02'", flevel='2'))
        optx.format_num_style(
            sheet1, 9, 1, other_receivables_end[0][0]+doubtful_debts_provision_o_end[0][0])

        # 期初存货
        inventories_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "(FNumber='1403' OR FNumber='1404')"))
        optx.format_num_style(
            sheet1, 10, 2, inventories_begin[0][0])
        # 期末存货
        inventories_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "(FNumber='1403' OR FNumber='1404')"))
        optx.format_num_style(
            sheet1, 10, 1, inventories_end[0][0])

        # 期初合同资产
        optx.format_num_style(
            sheet1, 11, 2, 0)
        # 期末合同资产
        optx.format_num_style(
            sheet1, 11, 1, 0)

        # 期初持有待售资产
        optx.format_num_style(
            sheet1, 12, 2, 0)
        # 期末持有待售资产
        optx.format_num_style(
            sheet1, 12, 1, 0)

        # 期初一年内到期的非流 动资产
        holding_the_investment_until_the_call_date_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1501'"))
        provision_for_diminution_in_value_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet(
                'Begin', db_name_needed[i], fyear, '1', "FNumber='1502'"))
        optx.format_num_style(
            sheet1, 13, 2, holding_the_investment_until_the_call_date_begin[0][0]-provision_for_diminution_in_value_begin[0][0])
        # 期末一年内到期的非流动资产
        holding_the_investment_until_the_call_date_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1501'"))
        provision_for_diminution_in_value_end = msg.ExecQuery(
            sqls.sqls_balance_sheet(
                'End', db_name_needed[i], fyear, fperiod, "FNumber='1502'"))
        optx.format_num_style(
            sheet1, 13, 1, holding_the_investment_until_the_call_date_end[0][0]-provision_for_diminution_in_value_end[0][0])

        # 期初其他流动资产
        other_current_assets_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1902'"))
        optx.format_num_style(
            sheet1, 14, 2, other_current_assets_begin[0][0])
        # 期末其他流动资产
        other_current_assets_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1902'"))
        optx.format_num_style(
            sheet1, 14, 1, other_current_assets_end[0][0])

        # 期初债权投资
        optx.format_num_style(
            sheet1, 17, 2, 0)
        # 期末债权投资
        optx.format_num_style(
            sheet1, 17, 1, 0)

        # 期初债权投资
        optx.format_num_style(
            sheet1, 18, 2, 0)
        # 期末债权投资
        optx.format_num_style(
            sheet1, 18, 1, 0)

        # 期初长期应收款
        long_term_receivables_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1503'"))
        optx.format_num_style(
            sheet1, 19, 2, long_term_receivables_begin[0][0])
        # 期末长期应收款
        long_term_receivables_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1503'"))
        optx.format_num_style(
            sheet1, 19, 1, long_term_receivables_end[0][0])

        # 期初长期股权投资
        long_term_equity_investments_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1511'"))
        optx.format_num_style(
            sheet1, 20, 2, long_term_equity_investments_begin[0][0])
        # 期末长期股权投资
        long_term_equity_investments_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1511'"))
        optx.format_num_style(
            sheet1, 20, 1, long_term_equity_investments_end[0][0])

        # 期初其他权益工具投资
        optx.format_num_style(
            sheet1, 21, 2, 0)
        # 期末其他权益工具投资
        optx.format_num_style(
            sheet1, 21, 1, 0)

        # 期初其他非流动金融资产
        optx.format_num_style(
            sheet1, 22, 2, 0)
        # 期末其他非流动金融资产
        optx.format_num_style(
            sheet1, 22, 1, 0)

        # 期初投资性房地产
        optx.format_num_style(
            sheet1, 23, 2, 0)
        # 期末投资性房地产
        optx.format_num_style(
            sheet1, 23, 1, 0)

        # 期初固定资产
        fixed_assets_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1601'"))
        accumulated_depreciation_f_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1602'"))
        optx.format_num_style(
            sheet1, 24, 2, fixed_assets_begin[0][0]+accumulated_depreciation_f_begin[0][0])
        # 期末固定资产
        fixed_assets_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1601'"))
        accumulated_depreciation_f_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1602'"))
        optx.format_num_style(
            sheet1, 24, 1, fixed_assets_end[0][0]+accumulated_depreciation_f_end[0][0])

        # 期初在建工程
        construction_in_progress_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1604'"))
        optx.format_num_style(
            sheet1, 25, 2, construction_in_progress_begin[0][0])
        # 期末在建工程
        construction_in_progress_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1604'"))
        optx.format_num_style(
            sheet1, 25, 1, construction_in_progress_end[0][0])

        # 期初生产性生物资产
        optx.format_num_style(
            sheet1, 26, 2, 0)
        # 期末生产性生物资产
        optx.format_num_style(
            sheet1, 26, 1, 0)

        # 期初油气资产
        optx.format_num_style(
            sheet1, 27, 2, 0)
        # 期末油气资产
        optx.format_num_style(
            sheet1, 27, 1, 0)

        # 期初无形资产
        intangible_assets_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1701'"))
        accumulated_depreciation_i_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1702'"))
        optx.format_num_style(
            sheet1, 28, 2, intangible_assets_begin[0][0]+accumulated_depreciation_i_begin[0][0])
        # 期末无形资产
        intangible_assets_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1701'"))
        accumulated_depreciation_i_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1702'"))
        optx.format_num_style(
            sheet1, 28, 1, intangible_assets_end[0][0]+accumulated_depreciation_i_end[0][0])

        # 期初开发支出
        optx.format_num_style(
            sheet1, 29, 2, 0)
        # 期末开发支出
        optx.format_num_style(
            sheet1, 29, 1, 0)

        # 期初商誉
        goodwill_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1711'"))
        optx.format_num_style(
            sheet1, 30, 2, goodwill_begin[0][0])
        # 期末商誉
        goodwill_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1711'"))
        optx.format_num_style(
            sheet1, 30, 1, goodwill_end[0][0])

        # 期初长期待摊费用
        long_term_prepaid_expenses_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1801'"))
        optx.format_num_style(
            sheet1, 31, 2, long_term_prepaid_expenses_begin[0][0])
        # 期末长期待摊费用
        long_term_prepaid_expenses_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1801'"))
        optx.format_num_style(
            sheet1, 31, 1, long_term_prepaid_expenses_end[0][0])

        # 期初递延所得税资产
        deferred_tax_assets_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='1811'"))
        optx.format_num_style(
            sheet1, 32, 2, deferred_tax_assets_begin[0][0])
        # 期末递延所得税资产
        deferred_tax_assets_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='1811'"))
        optx.format_num_style(
            sheet1, 32, 1, deferred_tax_assets_end[0][0])

        # 期初其他非流动资产
        optx.format_num_style(
            sheet1, 33, 2, 0)
        # 期末其他非流动资产
        optx.format_num_style(
            sheet1, 33, 1, 0)

        # 期初短期借款
        short_term_borrowings_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2001'"))
        optx.format_num_style(
            sheet1, 4, 5, -short_term_borrowings_begin[0][0])
        # 期末短期借款
        short_term_borrowings_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2001'"))
        optx.format_num_style(
            sheet1, 4, 4, -short_term_borrowings_end[0][0])

        # 期初交易性金融负债
        optx.format_num_style(
            sheet1, 5, 5, 0)
        # 期末交易性金融负债
        optx.format_num_style(
            sheet1, 5, 4, 0)

        # 期初衍生金融负债
        optx.format_num_style(
            sheet1, 6, 5, 0)
        # 期末衍生金融负债
        optx.format_num_style(
            sheet1, 6, 4, 0)

        # 期初应付票据及应付账款
        notes_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2201'"))
        accounts_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2202'"))
        optx.format_num_style(
            sheet1, 7, 5, -notes_payable_begin[0][0]-accounts_payable_begin[0][0])
        # 期末应付票据及应付账款
        notes_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2201'"))
        accounts_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2202'"))
        optx.format_num_style(
            sheet1, 7, 4, -notes_payable_end[0][0]-accounts_payable_end[0][0])

        # 期初预收账款
        advances_from_customers_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2203'"))
        optx.format_num_style(
            sheet1, 8, 5, -advances_from_customers_begin[0][0])
        # 期末预收账款
        advances_from_customers_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2203'"))
        optx.format_num_style(
            sheet1, 8, 4, -advances_from_customers_end[0][0])

        # 期初合同负债
        optx.format_num_style(
            sheet1, 9, 5, 0)
        # 期末合同负债
        optx.format_num_style(
            sheet1, 9, 4, 0)

        # 期初应付职工薪酬
        employee_benefits_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2211'"))
        optx.format_num_style(
            sheet1, 10, 5, -employee_benefits_payable_begin[0][0])
        # 期末应付职工薪酬
        employee_benefits_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2211'"))
        optx.format_num_style(
            sheet1, 10, 4, -employee_benefits_payable_end[0][0])

        # 期初应交税费
        taxes_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2221'"))
        optx.format_num_style(
            sheet1, 11, 5, -taxes_payable_begin[0][0])
        # 期末应交税费
        taxes_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2221'"))
        optx.format_num_style(
            sheet1, 11, 4, -taxes_payable_end[0][0])

        # 期初其他应付款
        other_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2241'"))
        optx.format_num_style(
            sheet1, 12, 5, -other_payable_begin[0][0])
        # 期末其他应付款
        other_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2241'"))
        optx.format_num_style(
            sheet1, 12, 4, -other_payable_end[0][0])

        # 期初 持有待售负债
        optx.format_num_style(
            sheet1, 13, 5, 0)
        # 期末 持有待售负债
        optx.format_num_style(
            sheet1, 13, 4, 0)

        # 期初一年内到期的非流动负债
        optx.format_num_style(
            sheet1, 14, 5, 0)
        # 期末一年内到期的非流动负债
        optx.format_num_style(
            sheet1, 14, 4, 0)

        # 期初其他流动负债
        optx.format_num_style(
            sheet1, 15, 5, 0)
        # 期末其他流动负债
        optx.format_num_style(
            sheet1, 15, 4, 0)

        # 期初长期借款
        long_term_borrowings_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2501'"))
        optx.format_num_style(
            sheet1, 18, 5, -long_term_borrowings_begin[0][0])
        # 期末长期借款
        long_term_borrowings_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2501'"))
        optx.format_num_style(
            sheet1, 18, 4, -long_term_borrowings_end[0][0])

        # 期初长期债券
        debentures_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2502'"))
        optx.format_num_style(
            sheet1, 19, 5, -debentures_payable_begin[0][0])
        # 期末长期债券
        debentures_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2502'"))
        optx.format_num_style(
            sheet1, 19, 4, -debentures_payable_end[0][0])

        # 期初优先债
        optx.format_num_style(
            sheet1, 20, 5, 0)
        # 期末优先债
        optx.format_num_style(
            sheet1, 20, 4, 0)

        # 期初永续债
        optx.format_num_style(
            sheet1, 21, 5, 0)
        # 期末永续债
        optx.format_num_style(
            sheet1, 21, 4, 0)

        # 期初长期应付款
        long_term_payable_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2701'"))
        optx.format_num_style(
            sheet1, 22, 5, -long_term_payable_begin[0][0])
        # 期末长期应付款
        long_term_payable_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2701'"))
        optx.format_num_style(
            sheet1, 22, 4, -long_term_payable_end[0][0])

        # 期初长期应付款
        provisions_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2801'"))
        optx.format_num_style(
            sheet1, 23, 5, -provisions_begin[0][0])
        # 期末长期应付款
        provisions_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2801'"))
        optx.format_num_style(
            sheet1, 23, 4, -provisions_end[0][0])

        # 期初递延收益
        deferred_income_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2401'"))
        optx.format_num_style(
            sheet1, 24, 5, -deferred_income_begin[0][0])
        # 期末递延收益
        deferred_income_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2401'"))
        optx.format_num_style(
            sheet1, 24, 4, -deferred_income_end[0][0])

        # 期初递延所得税负债
        deferred_tax_liabilities_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='2901'"))
        optx.format_num_style(
            sheet1, 25, 5, -deferred_tax_liabilities_begin[0][0])
        # 期末递延所得税负债
        deferred_tax_liabilities_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='2901'"))
        optx.format_num_style(
            sheet1, 25, 4, -deferred_tax_liabilities_end[0][0])

        # 期初其他非流动负债
        optx.format_num_style(
            sheet1, 26, 5, 0)
        # 期末其他非流动负债
        optx.format_num_style(
            sheet1, 26, 4, 0)

        # 期初股本
        paid_in_capital_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "(FNumber='4001' OR FNumber='4003')"))
        optx.format_num_style(
            sheet1, 30, 5, -paid_in_capital_begin[0][0])
        # 期末股本
        paid_in_capital_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "(FNumber='4001' OR FNumber='4003')"))
        optx.format_num_style(
            sheet1, 30, 4, -paid_in_capital_end[0][0])

        # 期初其他权益工具
        optx.format_num_style(
            sheet1, 31, 5, 0)
        # 期末其他权益工具
        optx.format_num_style(
            sheet1, 31, 4, 0)

        # 期初优先股
        optx.format_num_style(
            sheet1, 32, 5, 0)
        # 期末优先股
        optx.format_num_style(
            sheet1, 32, 4, 0)

        # 期初永续债
        optx.format_num_style(
            sheet1, 33, 5, 0)
        # 期末永续债
        optx.format_num_style(
            sheet1, 33, 4, 0)

        # 期初资本公积
        capital_surplus_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='4002'"))
        optx.format_num_style(
            sheet1, 34, 5, -capital_surplus_begin[0][0])
        # 期末资本公积
        capital_surplus_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='4002'"))
        optx.format_num_style(
            sheet1, 34, 4, -capital_surplus_end[0][0])

        # 期初库存股
        treasury_shares_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='4201'"))
        optx.format_num_style(
            sheet1, 35, 5, -treasury_shares_begin[0][0])
        # 期末库存股
        treasury_shares_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='4201'"))
        optx.format_num_style(
            sheet1, 35, 4, -treasury_shares_end[0][0])

        # 期初其他综合收益
        optx.format_num_style(
            sheet1, 36, 5, 0)
        # 期末其他综合收益
        optx.format_num_style(
            sheet1, 36, 4, 0)

        # 期初盈余公积
        surplus_reserve_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "FNumber='4101'"))
        optx.format_num_style(
            sheet1, 37, 5, -surplus_reserve_begin[0][0])
        # 期末盈余公积
        surplus_reserve_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "FNumber='4101'"))
        optx.format_num_style(
            sheet1, 37, 4, -surplus_reserve_end[0][0])

        # 期初未分配利润
        undistributed_profits_begin = msg.ExecQuery(
            sqls.sqls_balance_sheet('Begin', db_name_needed[i], fyear, '1', "(FNumber='4103' OR FNumber='4104')"))
        optx.format_num_style(
            sheet1, 38, 5, -undistributed_profits_begin[0][0])
        # 期末未分配利润
        undistributed_profits_end = msg.ExecQuery(
            sqls.sqls_balance_sheet('End', db_name_needed[i], fyear, fperiod, "(FNumber='4103' OR FNumber='4104')"))
        optx.format_num_style(
            sheet1, 38, 4, -undistributed_profits_end[0][0])

        # 利润表项目
        # 本期营业收入
        operating_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FCredit', db_name_needed[i], fyear, fperiod, "(FNumber = '6001' OR FNumber = '6051')"))
        optx.format_num_style(
            sheet2, 3, 1, operating_income_current[0][0])
        # 本年累计营业收入
        operating_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdCredit', db_name_needed[i], fyear, fperiod, "(FNumber = '6001' OR FNumber = '6051')"))
        optx.format_num_style(
            sheet2, 3, 2, operating_income_accumulated[0][0])

        # 本期营业成本
        operating_cost_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "(FNumber = '6401' OR FNumber = '6402')"))
        optx.format_num_style(
            sheet2, 4, 1, operating_cost_current[0][0])
        # 本年累计营业成本
        operating_cost_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "(FNumber = '6401' OR FNumber = '6402')"))
        optx.format_num_style(
            sheet2, 4, 2, operating_cost_accumulated[0][0])

        # 本期营业税金及附加
        business_tax_and_surcharges_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6403'"))
        optx.format_num_style(
            sheet2, 5, 1, business_tax_and_surcharges_current[0][0])
        # 本年累计营业税金及附加
        business_tax_and_surcharges_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6403'"))
        optx.format_num_style(
            sheet2, 5, 2, business_tax_and_surcharges_accumulated[0][0])

        # 本期销售费用
        sales_expense_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6601'"))
        optx.format_num_style(
            sheet2, 6, 1, sales_expense_current[0][0])
        # 本年累计销售费用
        sales_expense_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6601'"))
        optx.format_num_style(
            sheet2, 6, 2, sales_expense_accumulated[0][0])

        # 本期管理费用
        management_expense_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6602'"))
        optx.format_num_style(
            sheet2, 7, 1, management_expense_current[0][0])
        # 本年累计管理费用
        management_expense_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6602'"))
        optx.format_num_style(
            sheet2, 7, 2, management_expense_accumulated[0][0])

        # 本期研发费用
        r_d_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6605'"))
        optx.format_num_style(
            sheet2, 8, 1, r_d_current[0][0])
        # 本年累计研发费用
        r_d_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6605'"))
        optx.format_num_style(
            sheet2, 8, 2, r_d_accumulated[0][0])

        # 本期财务费用
        financial_expense_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603'"))
        optx.format_num_style(
            sheet2, 9, 1, financial_expense_current[0][0])
        # 本年累计财务费用
        financial_expense_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603'"))
        optx.format_num_style(
            sheet2, 9, 2, financial_expense_accumulated[0][0])

        # 本期财务费用-利息费用
        interest_expense_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603.02' AND Balance.FDetailID = 0", flevel='2'))
        optx.format_num_style(
            sheet2, 10, 1, interest_expense_current[0][0])
        # 本年累计财务费用-利息费用
        interest_expense_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603.02' AND Balance.FDetailID = 0", flevel='2'))
        optx.format_num_style(
            sheet2, 10, 2, interest_expense_accumulated[0][0])

        # 本期财务费用-利息收入
        interest_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603.03' AND Balance.FDetailID = 0", flevel='2'))
        optx.format_num_style(
            sheet2, 11, 1, interest_income_current[0][0])
        # 本年累计财务费用-利息收入
        interest_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6603.03' AND Balance.FDetailID = 0", flevel='2'))
        optx.format_num_style(
            sheet2, 11, 2, interest_income_accumulated[0][0])

        # 本期资产减值损失
        asset_impairment_loss_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6701'"))
        optx.format_num_style(
            sheet2, 12, 1, asset_impairment_loss_current[0][0])
        # 本年累计资产减值损失
        asset_impairment_loss_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6701'"))
        optx.format_num_style(
            sheet2, 12, 2, asset_impairment_loss_accumulated[0][0])

        # 本期信用减值损失
        optx.format_num_style(
            sheet2, 13, 1, 0)
        # 本年累计信用减值损失
        optx.format_num_style(
            sheet2, 13, 2, 0)

        # 本期其他收益
        other_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6117'"))
        optx.format_num_style(
            sheet2, 14, 1, other_income_current[0][0])
        # 本年累计其他收益
        other_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6117'"))
        optx.format_num_style(
            sheet2, 14, 2, other_income_accumulated[0][0])

        # 本期投资收益
        incestment_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6111'"))
        optx.format_num_style(
            sheet2, 15, 1, incestment_income_current[0][0])
        # 本年累计投资收益
        incestment_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6111'"))
        optx.format_num_style(
            sheet2, 15, 2, incestment_income_accumulated[0][0])

        # 本期对联营企业与合营企业的投资收益
        optx.format_num_style(
            sheet2, 16, 1, 0)
        # 本年累计对联营企业与合营企业的投资收益
        optx.format_num_style(
            sheet2, 16, 2, 0)

        # 本期 净敞口套期收益
        optx.format_num_style(
            sheet2, 17, 1, 0)
        # 本年累计 净敞口套期收益
        optx.format_num_style(
            sheet2, 17, 2, 0)

        # 本期公允价值变动收益
        fair_value_change_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6111'"))
        optx.format_num_style(
            sheet2, 18, 1, fair_value_change_income_current[0][0])
        # 本年累计公允价值变动收益
        fair_value_change_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6111'"))
        optx.format_num_style(
            sheet2, 18, 2, fair_value_change_income_accumulated[0][0])

        # 本期资产处置收益
        assets_disposal_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6115'"))
        optx.format_num_style(
            sheet2, 19, 1, assets_disposal_income_current[0][0])
        # 本年累计资产处置收益
        assets_disposal_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6115'"))
        optx.format_num_style(
            sheet2, 19, 2, assets_disposal_income_accumulated[0][0])

        # 本期营业外支出
        operating_income_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6301'"))
        optx.format_num_style(
            sheet2, 21, 1, operating_income_current[0][0])
        # 本年累计营业外支出
        operating_income_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6301'"))
        optx.format_num_style(
            sheet2, 21, 2, operating_income_accumulated[0][0])

        # 本期营业外收入
        operating_cost_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6711'"))
        optx.format_num_style(
            sheet2, 22, 1, operating_cost_current[0][0])
        # 本年累计营业外收入
        operating_cost_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6711'"))
        optx.format_num_style(
            sheet2, 22, 2, operating_cost_accumulated[0][0])

        # 本期所得税费用
        income_tax_expense_current = msg.ExecQuery(
            sqls.sqls_profit_statement('FDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6801'"))
        optx.format_num_style(
            sheet2, 24, 1, income_tax_expense_current[0][0])
        # 本年累计所得税费用
        income_tax_expense_accumulated = msg.ExecQuery(
            sqls.sqls_profit_statement('FYtdDebit', db_name_needed[i], fyear, fperiod, "FNumber = '6801'"))
        optx.format_num_style(
            sheet2, 24, 2, income_tax_expense_accumulated[0][0])

        book.save('财务报表_'+fyear+'年'+fperiod+'月_'+company_name_needed[i]+'.xls')
