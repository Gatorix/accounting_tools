--货币资金_期初
SELECT ISNULL(SUM(FBeginBalance),0) AS 期初余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
    AND (FNumber=1001 OR FNumber=1002 OR FNumber=1012)

--货币资金_期末
SELECT ISNULL(SUM(FEndBalance),0) AS 期末余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=12 AND FLevel=1 AND Balance.FCurrencyID=1
    AND (FNumber=1001 OR FNumber=1002 OR FNumber=1012)

--交易性金融资产_期初
SELECT ISNULL(SUM(FBeginBalance),0) AS 期初余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
    AND FNumber=1101

--交易性金融资产_期末
SELECT ISNULL(SUM(FEndBalance),0) AS 期末余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
    AND FNumber=1101

--应收票据_期初
SELECT ISNULL(SUM(FBeginBalance),0) AS 期初余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
    AND FNumber=1121

--应收票据_期末
SELECT ISNULL(SUM(FEndBalance),0) AS 期末余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel=1 AND Balance.FCurrencyID=1
    AND FNumber=1121