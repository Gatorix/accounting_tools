
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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance >= 0
        AND Balance.FEndBalance >= 0)
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'

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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance < 0
        AND Balance.FEndBalance < 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'

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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance > 0
        AND Balance.FEndBalance > 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'

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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance <= 0
        AND Balance.FEndBalance <= 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'
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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance <= 0
        AND Balance.FEndBalance > 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'
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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = -1
        AND Balance.FBeginBalance > 0
        AND Balance.FEndBalance <= 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'

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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance >= 0
        AND Balance.FEndBalance < 0 )
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'
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
        AIS20160105001.dbo.t_Balance AS Balance
        INNER JOIN AIS20160105001.dbo.t_Account AS Account
        ON Account.FAccountID=Balance.FAccountID)
        INNER JOIN AIS20160105001.dbo.t_ItemDetailV AS ItemDetailV
        ON Balance.FDetailID=ItemDetailV.FDetailID)
        INNER JOIN AIS20160105001.dbo.t_Item AS Item
        ON ItemDetailV.FItemID=Item.FItemID)

    WHERE (Account.FDC = 1
        AND Balance.FBeginBalance < 0
        AND Balance.FEndBalance >= 0)
        AND Balance.FYear=2018
        AND Balance.FPeriod=12
        AND Balance.FCurrencyID=1
        AND Account.FLevel>=1
        AND Account.FLevel<=10
        AND (Balance.FBeginBalance<>0
        OR Balance.FDebit<>0
        OR Balance.FCredit<>0
        OR Balance.FYtdDebit<>0
        OR Balance.FYtdCredit<>0
        OR Balance.FEndBalance<>0)
        AND Account.FNumber>='1001'
        AND Account.FNumber<='9999'


ORDER BY 科目代码 , 核算项目名称