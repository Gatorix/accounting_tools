SELECT FBeginBalance AS 期初余额
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel<=2 AND Balance.FCurrencyID=1
    AND FNumber=1231.01 AND FNumber=1122


    --ISNULL(SUM(FBeginBalance),0) AS 期初余额


SELECT *
FROM AIS20160105001.dbo.t_Balance Balance
    INNER JOIN AIS20160105001.dbo.t_Account Account
    ON Account.FAccountID=Balance.FAccountID
WHERE FYear=2018 AND FPeriod=1 AND FLevel<=2 AND Balance.FCurrencyID=1 AND FNumber=1231.01
ORDER BY FNumber