import xlwt

print('写入模板参数表...')

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


workbook = xlwt.Workbook()
sheet2 = workbook.add_sheet('t_Schema', cell_overwrite_ok=True)

valuerow = len(value)
for row in range(valuerow):
    for col in range(0, len(value[row])):
        sheet2.write(row, col, value[row][col])


workbook.save("123.xls")
