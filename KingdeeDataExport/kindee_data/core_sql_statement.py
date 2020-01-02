def sql_balance(db, year, period, fnumber_begin, fnumber_end, minlevel):
    '''
    db, year, period, minlevel='10', maxlevel='1'
    通过用户输入的参数，从数据库获取完整的科目余额
    year代表会计年度，period代表期间，maxlevel代表最高级别科目，minlevel代表最末级，currencyid为货币代码。
    '''
    return '''

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        ( + Balance.FBeginBalance ) AS 期初借方余额,
        0 AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        ( + Balance.FEndBalance ) AS 期末借方余额,
        0 AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance >= 0
        AND Balance.FEndBalance >= 0)
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''

UNION ALL

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        0 AS 期初借方余额,
        ( - Balance.FBeginBalance ) AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        0 AS 期末借方余额,
        ( - Balance.FEndBalance ) AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance < 0
        AND Balance.FEndBalance < 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''

UNION ALL

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        ( + Balance.FBeginBalance ) AS 期初借方余额,
        0 AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        ( + Balance.FEndBalance ) AS 期末借方余额,
        0 AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance > 0
        AND Balance.FEndBalance > 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''

UNION ALL

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        0 AS 期初借方余额,
        ( - Balance.FBeginBalance ) AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        0 AS 期末借方余额,
        ( - Balance.FEndBalance ) AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance <= 0
        AND Balance.FEndBalance <= 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''
UNION ALL

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        0 AS 期初借方余额,
        ( - Balance.FBeginBalance ) AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        ( + Balance.FEndBalance ) AS 期末借方余额,
        0 AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance <= 0
        AND Balance.FEndBalance > 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''
UNION ALL
    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        ( + Balance.FBeginBalance ) AS 期初借方余额,
        0 AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        0 AS 期末借方余额,
        ( - Balance.FEndBalance ) AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance > 0
        AND Balance.FEndBalance <= 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''

UNION ALL

    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        ( + Balance.FBeginBalance ) AS 期初借方余额,
        0 AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        0 AS 期末借方余额,
        ( - Balance.FEndBalance ) AS 期末贷方余额


    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance >= 0
        AND Balance.FEndBalance < 0 )
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''
UNION ALL
    SELECT
        Account.FNumber AS 科目代码,
        CONVERT(NVARCHAR(50),Account.FFullName) AS 科目名称,
        CONVERT(NVARCHAR(50),Item.FName) AS 核算项目名称,
        0 AS 期初借方余额,
        ( - Balance.FBeginBalance ) AS 期初贷方余额,
        Balance.FDebit AS 本期借方发生额,
        Balance.FCredit AS 本期贷方发生额,
        Balance.FYtdDebit AS 本年累计借方发生额,
        Balance.FYtdCredit AS 本年累计贷方发生额,
        ( + Balance.FEndBalance ) AS 期末借方余额,
        0 AS 期末贷方余额

    FROM (((
        '''+db+'''.dbo.t_Balance AS Balance
        INNER JOIN '''+db+'''.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN '''+db+'''.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN '''+db+'''.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance < 0
        AND Balance.FEndBalance >= 0)
        AND Balance.FYear='''+year+'''
        AND Balance.FPeriod='''+period+'''
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<='''+minlevel+'''
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='''+fnumber_begin+'''
        AND Account.FNumber<='''+fnumber_end+'''


ORDER BY 科目代码 , 核算项目名称

    '''


# def sql_statement_cash_at_bank_and_on_hand_begin(db, FYear):
#     # 货币资金_期初
#     return '''
# SELECT ISNULL(SUM(FBeginBalance),0) AS 期初余额
# FROM '''+db+'''.dbo.t_Balance Balance
#     INNER JOIN '''+db+'''.dbo.t_Account Account
#     ON Account.FAccountID=Balance.FAccountID
# WHERE FYear='''+FYear+''' AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
#     AND (FNumber=1001 OR FNumber=1002 OR FNumber=1012)
# '''


def sqls_balance_sheet(eb, db, FYear, FPeriod, statement, flevel='1'):
    # 资产负债表项目
    return '''
SELECT ISNULL(SUM(F'''+eb+'''Balance),0)
FROM '''+db+'''.dbo.t_Balance Balance
    INNER JOIN '''+db+'''.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear='''+FYear+''' AND FPeriod='''+FPeriod+''' AND FLevel='''+flevel+''' AND Balance.FCurrencyID=1
    AND '''+statement+'''
'''


def sqls_profit_statement(cd, db, year, FPeriod, statement, flevel='1'):
    # 利润表项目
    return '''
SELECT ISNULL(SUM('''+cd+'''),0)
FROM '''+db+'''.dbo.t_Balance Balance
    INNER JOIN '''+db+'''.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear='''+year+''' AND FPeriod='''+FPeriod+''' AND FLevel='''+flevel+''' AND Balance.FCurrencyID=1
    AND '''+statement+'''
'''
