import xlwt


def create_book():
    return xlwt.Workbook()


def get_date(year, mon):
    if mon == 1 or mon == 3 or mon == 5 or mon == 7 or mon == 8 or mon == 10 or mon == 12:
        return 31
    elif mon == 4 or mon == 6 or mon == 9 or mon == 11:
        return 30
    else:
        if year % 4 == 0 and mon == 2:
            return 29
        if year % 4 != 0 and mon == 2:
            return 28


def write_header_line_trail_balance(sheet):
    header_line = ['科目代码', '科目名称', '核算项目名称', '期初借方余额', '期初贷方余额',
                   '本期借方发生额', '本期贷方发生额', '本年累计借方发生额',
                   '本年累计贷方发生额', '期末借方余额', '期末贷方余额']
    for i in range(len(header_line)):
        sheet.write(0, i, header_line[i])


def write_db_result_trail_balance(sheet, result):
    for row in range(len(result)):
        for col in range(0, len(result[row])):
            sheet.write(row+1, col, result[row][col])


def write_balance_sheet_fixed(sheet, corp_name, year, period):
    fixed_value_balance_sheet = [
        ['资产', '流动资产：', '  货币资金', '  交易性金融资产', '  衍生金融资产',
         '  应收票据及应收账款', '  预付账款', '  其他应收款', '  存货', '  合同资产', '  持有待售资产', '  一年内到期的非流动资产',
         '  其他流动资产', '流动资产合计', '非流动资产：', '  债权投资',
         '  其他债权投资', '  长期应收款', '  长期股权投资', '  其他权益工具投资', '  其他非流动金融资产',
         '  投资性房地产', '  固定资产', '  在建工程', '  生产性生物资产', '  油气资产', '  无形资产', '  开发支出',
         '  商誉', '  长期待摊费用', '  递延所得税资产', '  其他非流动资产', '非流动资产合计', '', '', '', '', '', '资产总计'],
        ['期末余额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
            '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['年初余额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
            '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['负债和所有者权益（或股东权益）', '流动负债：', '  短期借款', '  交易性金融负债', '  衍生金融负债',
         '  应付票据及应付账款', '  预收账款', '  合同负债', '  应付职工薪酬', '  应交税费', '  其他应付款', '  持有待售负债',
         '  一年内到期的非流动负债', '  其他流动负债', '流动负债合计', '非流动负债：', '  长期借款', '  应付债券',
         '  其中：优先债', '        永续债',
         '  长期应付款', '  预计负债', '  递延收益', '  递延所得税负债', '  其他非流动负债', '非流动负债合计', '负债合计',
         '所有者权益（或股东权益）：', '  实收资本（或股本）', '  其他权益工具', '  其中：优先股', '        永续债',
         '  资本公积', '    减：库存股', '  其他综合收益', '  盈余公积', '  未分配利润', '所有者权益（或股东权益）合计',
         '负债和所有者权益（或股东权益）总计'],
        ['期末余额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
            '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['年初余额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
            '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
    ]

    alignment_center = xlwt.Alignment()  # Create Alignment
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_center.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 11*20
    style = xlwt.XFStyle()
    style.borders = borders
    style.font = font
    style.alignment = alignment_center

    for c in range(len(fixed_value_balance_sheet)):
        for r in range(0, len(fixed_value_balance_sheet[c])):
            sheet.write(r+2, c, fixed_value_balance_sheet[c][r], style)
            sheet.row(r+2).height_mismatch = True
            sheet.row(r+2).height = 400

    # top_row = 0
    # bottom_row = 0
    # left_column = 0
    # right_column = 1
    # sheet.write_merge(top_row, bottom_row, left_column, right_column, 'Long Cell')
    alignment_center = xlwt.Alignment()  # Create Alignment
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment_center.horz = xlwt.Alignment.HORZ_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_center.vert = xlwt.Alignment.VERT_CENTER

    alignment_left = xlwt.Alignment()
    alignment_left.horz = xlwt.Alignment.HORZ_LEFT
    alignment_left.vert = xlwt.Alignment.VERT_CENTER

    alignment_right = xlwt.Alignment()
    alignment_right.horz = xlwt.Alignment.HORZ_RIGHT
    alignment_right.vert = xlwt.Alignment.VERT_CENTER

    font_all = xlwt.Font()
    font_all.name = '宋体'
    font_all.height = 11*20

    style_date = xlwt.XFStyle()
    style_date.alignment = alignment_center
    style_date.font = font_all
    sheet.write_merge(1, 1, 2, 3, year+'年'+period+'月' +
                      str(get_date(int(year), int(period)))+'日', style_date)

    style_dw = xlwt.XFStyle()
    style_dw.alignment = alignment_left
    style_dw.font = font_all
    sheet.write_merge(1, 1, 0, 1, '编制单位：'+corp_name, style_dw)

    style_y = xlwt.XFStyle()
    style_y.alignment = alignment_right
    style_y.font = font_all
    sheet.write(1, 5, '单位：元', style_y)

    # 标题行样式
    font_head = xlwt.Font()
    font_head.bold = True
    font_head.name = '宋体'
    font_head.height = 18*20
    alignment_head = xlwt.Alignment()  # Create Alignment
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment_head.horz = xlwt.Alignment.HORZ_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_head.vert = xlwt.Alignment.VERT_CENTER
    style_head = xlwt.XFStyle()  # Create Style
    style_head.alignment = alignment_head  # Add Alignment to Style
    style_head.font = font_head

    sheet.write_merge(0, 0, 0, 5, '资产负债表', style_head)

    style_num = xlwt.XFStyle()

    borders_num = xlwt.Borders()
    borders_num.left = xlwt.Borders.THIN
    borders_num.right = xlwt.Borders.THIN
    borders_num.top = xlwt.Borders.THIN
    borders_num.bottom = xlwt.Borders.THIN
    borders_num.left_colour = 0x40
    borders_num.right_colour = 0x40
    borders_num.top_colour = 0x40
    borders_num.bottom_colour = 0x40

    style_num.borders = borders_num

    alignment_num = xlwt.Alignment()  # Create Alignment
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_num.vert = xlwt.Alignment.VERT_CENTER

    style_num.num_format_str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

    style_num.alignment = alignment_num

    style_num.font = font_all

    sheet.write(15, 1, xlwt.Formula('SUM(B5: B15)'), style_num)

    sheet.write(15, 2, xlwt.Formula('SUM(C5: C15)'), style_num)

    sheet.write(34, 1, xlwt.Formula('SUM(B18:B34)'), style_num)

    sheet.write(34, 2, xlwt.Formula('SUM(C18:C34)'), style_num)

    sheet.write(40, 1, xlwt.Formula('B35+B16'), style_num)

    sheet.write(40, 2, xlwt.Formula('C35+C16'), style_num)

    sheet.write(16, 4, xlwt.Formula('SUM(E5:E16)'), style_num)

    sheet.write(16, 5, xlwt.Formula('SUM(F5:F16)'), style_num)

    sheet.write(27, 4, xlwt.Formula('SUM(E19:E20)+SUM(E23:E27)'), style_num)

    sheet.write(27, 5, xlwt.Formula('SUM(F19:F20)+SUM(F23:F27)'), style_num)

    sheet.write(28, 4, xlwt.Formula('E17+E28'), style_num)

    sheet.write(28, 5, xlwt.Formula('F17+F28'), style_num)

    sheet.write(39, 4, xlwt.Formula('SUM(E31:E39)'), style_num)

    sheet.write(39, 5, xlwt.Formula('SUM(F31:F39)'), style_num)

    sheet.write(40, 4, xlwt.Formula('E40+E29'), style_num)

    sheet.write(40, 5, xlwt.Formula('F40+F29'), style_num)
    for i in range(6):
        sheet.col(i).width = 256*25

    # for i in range(41):
    #     sheet.row(i).height_mismatch = True
    #     sheet.row(i).height = 400
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 600
    sheet.row(1).height_mismatch = True
    sheet.row(1).height = 400


def write_profit_statement_fixed(sheet, corp_name, year, period):
    fixed_value_profit_statement = [['项目',
                                     '一、营业收入', '  减：营业成本', '      营业税金及附加', '      销售费用', '      管理费用',
                                     '      研发费用', '      财务费用', '        其中：利息费用', '              利息收入',
                                     '      资产减值损失', '      信用减值损失', '  加：其他收益', '      投资收益',
                                     '        其中：对联营企业与合营企业的投资收益', '      净敞口套期收益', '      公允价值变动收益',
                                     '      资产处置收益', '二、营业利润', '    加: 营业外收入', '    减：营业外支出',
                                     '三、利润总额', '    减：所得税费用', '四、净利润', '  （一）持续经营利润', '  （二）终止经营利润',
                                     '五、其他综合收益的税后净额', '  （一）不能重分类进损益的其他综合收益', '     1.重新计量设定收益计划变动额',
                                     '     2.权益法下不能转损益的其他综合收益', '     3.其他权益工具投资公允价值变动', '     4.企业自身信用风险公允价值变动',
                                     '      ……', '  （二）将重分类进损益的其他综合收益', '     1．权益法下可转损益的其他综合收益', '     2．其它债权投资公允价值变动',
                                     '     3．金融资产重分类计入其他综合收益的金额', '     4.其他债权投资信用减值准备', '     5.现金流量套期准备',
                                     '     6.外币报表折算差额', '     ……', '七、每股收益：', '  基本每股收益', '  稀释每股收益'],
                                    ['本期金额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                     '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                                    ['本年累计金额', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                     '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']]

    # book = xlwt.Workbook()
    # # write_balance_sheet_fixed(book)
    # sheet = book.add_sheet('利润表', True)
    alignment_center = xlwt.Alignment()  # Create Alignment
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_center.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 11*20
    style = xlwt.XFStyle()
    style.borders = borders
    style.font = font
    style.alignment = alignment_center

    for c in range(len(fixed_value_profit_statement)):
        for r in range(0, len(fixed_value_profit_statement[c])):
            sheet.write(r+2, c, fixed_value_profit_statement[c][r], style)
            sheet.row(r+2).height_mismatch = True
            sheet.row(r+2).height = 400

    alignment_center = xlwt.Alignment()  # Create Alignment
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment_center.horz = xlwt.Alignment.HORZ_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_center.vert = xlwt.Alignment.VERT_CENTER

    alignment_left = xlwt.Alignment()
    alignment_left.horz = xlwt.Alignment.HORZ_LEFT
    alignment_left.vert = xlwt.Alignment.VERT_CENTER

    alignment_right = xlwt.Alignment()
    alignment_right.horz = xlwt.Alignment.HORZ_RIGHT
    alignment_right.vert = xlwt.Alignment.VERT_CENTER

    font_all = xlwt.Font()
    font_all.name = '宋体'
    font_all.height = 11*20

    style_date = xlwt.XFStyle()
    style_date.alignment = alignment_center
    style_date.font = font_all
    sheet.write(1, 1, year+'年'+period+'月', style_date)

    style_dw = xlwt.XFStyle()
    style_dw.alignment = alignment_left
    style_dw.font = font_all
    sheet.write(1, 0, '编制单位：'+corp_name, style_dw)

    style_y = xlwt.XFStyle()
    style_y.alignment = alignment_right
    style_y.font = font_all
    sheet.write(1, 2, '单位：元', style_y)

    font_head = xlwt.Font()
    font_head.bold = True
    font_head.name = '宋体'
    font_head.height = 18*20
    alignment_head = xlwt.Alignment()  # Create Alignment
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment_head.horz = xlwt.Alignment.HORZ_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_head.vert = xlwt.Alignment.VERT_CENTER
    style_head = xlwt.XFStyle()  # Create Style
    style_head.alignment = alignment_head  # Add Alignment to Style
    style_head.font = font_head

    sheet.write_merge(0, 0, 0, 2, '利润表', style_head)

    style_num = xlwt.XFStyle()

    borders_num = xlwt.Borders()
    borders_num.left = xlwt.Borders.THIN
    borders_num.right = xlwt.Borders.THIN
    borders_num.top = xlwt.Borders.THIN
    borders_num.bottom = xlwt.Borders.THIN
    borders_num.left_colour = 0x40
    borders_num.right_colour = 0x40
    borders_num.top_colour = 0x40
    borders_num.bottom_colour = 0x40

    style_num.borders = borders_num

    alignment_num = xlwt.Alignment()  # Create Alignment
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_num.vert = xlwt.Alignment.VERT_CENTER

    style_num.num_format_str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

    style_num.alignment = alignment_num

    style_num.font = font_all

    sheet.write(20, 1, xlwt.Formula(
        'B4-SUM(B5:B10)-SUM(B13:B14)+SUM(B15:B16)+SUM(B18:B20)'), style_num)

    sheet.write(20, 2, xlwt.Formula(
        'C4-SUM(C5:C10)-SUM(C13:C14)+SUM(C15:C16)+SUM(C18:C20)'), style_num)

    sheet.write(23, 1, xlwt.Formula('B21+B22-B23'), style_num)

    sheet.write(23, 2, xlwt.Formula('C21+C22-C23'), style_num)

    sheet.write(25, 1, xlwt.Formula('B24-B25'), style_num)

    sheet.write(25, 2, xlwt.Formula('C24-C25'), style_num)

    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 600
    sheet.row(1).height_mismatch = True
    sheet.row(1).height = 400
    # for i in range(3):
    sheet.col(0).width = 256*50
    sheet.col(1).width = 256*25
    sheet.col(2).width = 256*25


def format_num_style(sheet, row, col, value):
    style_num = xlwt.XFStyle()

    font_all = xlwt.Font()
    font_all.name = '宋体'
    font_all.height = 11*20

    borders_num = xlwt.Borders()
    borders_num.left = xlwt.Borders.THIN
    borders_num.right = xlwt.Borders.THIN
    borders_num.top = xlwt.Borders.THIN
    borders_num.bottom = xlwt.Borders.THIN
    borders_num.left_colour = 0x40
    borders_num.right_colour = 0x40
    borders_num.top_colour = 0x40
    borders_num.bottom_colour = 0x40

    alignment_num = xlwt.Alignment()  # Create Alignment
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_num.vert = xlwt.Alignment.VERT_CENTER

    style_num.borders = borders_num

    style_num.num_format_str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

    style_num.alignment = alignment_num

    style_num.font = font_all

    sheet.write(row, col, value, style_num)
