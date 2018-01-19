Sub continuationLoop()

Dim rng1 As Range
Dim rng2 As Range
Dim batch As String

Dim tickerRange As Range
Dim x As Integer
Dim rng As Range
Dim tickers() As Variant
ReDim tickers(1 To 100) As Variant

Set tickerRange = Application.InputBox(prompt:="Select tickers", Type:=8)
x = tickerRange.Cells.Count
Set rng1 = tickerRange.Cells(1, 1)
Set rng2 = tickerRange.Cells(1, 1)

MsgBox x

If x >= 100 Then
y = x / 100
Z = x Mod 100
j = 1
While j <= y
    For i = 1 To 100 Step 1
        rng1.Select
        tickers(i) = Selection.Value
        rng1.Offset(1, 0).Select
        Set rng1 = ActiveCell
    Next
    batch = Join(tickers, ",")
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", "https://api.iextrading.com/1.0/stock/market/batch?symbols=" & batch & "&types=company,quote,stats,financials,earnings,dividends"
    MyRequest.Send
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    Dim companyName, exchange, sector, industry, CEO, issueType, dividendType As Variant
    Dim latestPrice, openPrice, closePrice, low, high, change, changePercent, latestVolume, avgTotalVolume, week52Low, week52High, day50MovingAvg, day200MovingAvg, day5ChangePercent, month1ChangePercent, month3ChangePercent, month6ChangePercent, ytdChangePercent, year1ChangePercent, year3ChangePercent, year5ChangePercent, beta, marketcap, sharesOutstanding, float, revenue, revenuePerShare, revenuePerEmployee, EBITDA, grossProfit, profitMargin, cash, debt, returnOnEquity, returnOnAssets, returnOnCapital, peRatio, peRatioLow, peRatioHigh, priceToSales, priceToBook, shortRatio, costOfRevenue, opeartingRevenue, totalRevenue, opeartingIncome, netIncome, researchAndDevelopment, opeartingExpenses, currentAssets, totalAssets, totalLiabilities, currentCash, currentDebt, totalCash, totalDebt, shareholderEquity, cashChange, cashFlow, operatingGainsLosses, amount, dividendRate, dividendYield As Variant
    Dim exDate, paymentDate, declaredDate, recordDate As Variant
    Dim Dict As New Dictionary
    Dict.CompareMode = CompareMethod.TextCompare
    MsgBox batch
    For i = 1 To 100 Step 1
        Dict("A") = rng2.Value
        companyName = Json(Dict.Item("A"))("company")("companyName")
        exchange = Json(Dict.Item("A"))("company")("exchange")
        sector = Json(Dict.Item("A"))("company")("sector")
        industry = Json(Dict.Item("A"))("company")("industry")
        CEO = Json(Dict.Item("A"))("company")("CEO")
        issueType = Json(Dict.Item("A"))("company")("issueType")
        latestPrice = Json(Dict.Item("A"))("quote")("latestPrice")
        openPrice = Json(Dict.Item("A"))("quote")("open")
        closePrice = Json(Dict.Item("A"))("quote")("close")
        low = Json(Dict.Item("A"))("quote")("low")
        high = Json(Dict.Item("A"))("quote")("high")
        change = Json(Dict.Item("A"))("quote")("change")
        changePercent = Json(Dict.Item("A"))("quote")("changePercent")
        latestVolume = Json(Dict.Item("A"))("quote")("latestVolume")
        avgTotalVolume = Json(Dict.Item("A"))("quote")("avgTotalVolume")
        week52Low = Json(Dict.Item("A"))("quote")("week52Low")
        week52High = Json(Dict.Item("A"))("quote")("week52High")
        day50MovingAvg = Json(Dict.Item("A"))("stats")("day50MovingAvg")
        day200MovingAvg = Json(Dict.Item("A"))("stats")("day200MovingAvg")
        day5ChangePercent = Json(Dict.Item("A"))("stats")("day5ChangePercent")
        month1ChangePercent = Json(Dict.Item("A"))("stats")("month1ChangePercent")
        month3ChangePercent = Json(Dict.Item("A"))("stats")("month3ChangePercent")
        month6ChangePercent = Json(Dict.Item("A"))("stats")("month6ChangePercent")
        ytdChangePercent = Json(Dict.Item("A"))("stats")("ytdChangePercent")
        year1ChangePercent = Json(Dict.Item("A"))("stats")("year1ChangePercent")
        year3ChangePercent = Json(Dict.Item("A"))("stats")("year3ChangePercent")
        year5ChangePercent = Json(Dict.Item("A"))("stats")("year5ChangePercent")
        beta = Json(Dict.Item("A"))("stats")("beta")
        marketcap = Json(Dict.Item("A"))("stats")("marketcap")
        sharesOutstanding = Json(Dict.Item("A"))("stats")("sharesOutstanding")
        float = Json(Dict.Item("A"))("stats")("float")
        revenue = Json(Dict.Item("A"))("stats")("revenue")
        revenuePerShare = Json(Dict.Item("A"))("stats")("revenuePerShare")
        revenuePerEmployee = Json(Dict.Item("A"))("stats")("revenuePerEmployee")
        EBITDA = Json(Dict.Item("A"))("stats")("EBITDA")
        grossProfit = Json(Dict.Item("A"))("stats")("grossProfit")
        profitMargin = Json(Dict.Item("A"))("stats")("profitMargin")
        cash = Json(Dict.Item("A"))("stats")("cash")
        debt = Json(Dict.Item("A"))("stats")("debt")
        returnOnEquity = Json(Dict.Item("A"))("stats")("returnOnEquity")
        returnOnAssets = Json(Dict.Item("A"))("stats")("returnOnAssets")
        returnOnCapital = Json(Dict.Item("A"))("stats")("returnOnCapital")
        peRatio = Json(Dict.Item("A"))("quote")("peRatio")
        peRatioLow = Json(Dict.Item("A"))("stats")("peRatioLow")
        peRatioHigh = Json(Dict.Item("A"))("stats")("peRatioHigh")
        priceToSales = Json(Dict.Item("A"))("stats")("priceToSales")
        priceToBook = Json(Dict.Item("A"))("stats")("priceToBook")
        shortRatio = Json(Dict.Item("A"))("stats")("shortRatio")
        grossProfit = Json(Dict.Item("A"))("stats")("grossProfit")
        costOfRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("costOfRevenue")
        opeartingRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingRevenue")
        totalRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("totalRevenue")
        opeartingIncome = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingIncome")
        netIncome = Json(Dict.Item("A"))("financials")("financials")(1)("netIncome")
        researchAndDevelopment = Json(Dict.Item("A"))("financials")("financials")(1)("researchAndDevelopment")
        opeartingExpenses = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingExpenses")
        currentAssets = Json(Dict.Item("A"))("financials")("financials")(1)("currentAssets")
        totalAssets = Json(Dict.Item("A"))("financials")("financials")(1)("totalAssets")
        totalLiabilities = Json(Dict.Item("A"))("financials")("financials")(1)("totalLiabilities")
        currentCash = Json(Dict.Item("A"))("financials")("financials")(1)("currentCash")
        currentDebt = Json(Dict.Item("A"))("financials")("financials")(1)("currentDebt")
        totalCash = Json(Dict.Item("A"))("financials")("financials")(1)("totalCash")
        totalDebt = Json(Dict.Item("A"))("financials")("financials")(1)("totalDebt")
        shareholderEquity = Json(Dict.Item("A"))("financials")("financials")(1)("shareholderEquity")
        cashChange = Json(Dict.Item("A"))("financials")("financials")(1)("cashChange")
        cashFlow = Json(Dict.Item("A"))("financials")("financials")(1)("cashFlow")
        operatingGainsLosses = Json(Dict.Item("A"))("financials")("financials")(1)("operatingGainsLosses")
        rng2.Offset(0, 1).Value = companyName
        rng2.Offset(0, 2).Value = exchange
        rng2.Offset(0, 3).Value = sector
        rng2.Offset(0, 4).Value = industry
        rng2.Offset(0, 5).Value = CEO
        rng2.Offset(0, 6).Value = issueType
        rng2.Offset(0, 7).Value = latestPrice
        rng2.Offset(0, 8).Value = openPrice
        rng2.Offset(0, 9).Value = closePrice
        rng2.Offset(0, 10).Value = low
        rng2.Offset(0, 11).Value = high
        rng2.Offset(0, 12).Value = change
        rng2.Offset(0, 13).Value = changePercent
        rng2.Offset(0, 14).Value = latestVolume
        rng2.Offset(0, 15).Value = avgTotalVolume
        rng2.Offset(0, 16).Value = week52Low
        rng2.Offset(0, 17).Value = week52High
        rng2.Offset(0, 18).Value = day50MovingAvg
        rng2.Offset(0, 19).Value = day200MovingAvg
        rng2.Offset(0, 20).Value = day5ChangePercent
        rng2.Offset(0, 21).Value = month1ChangePercent
        rng2.Offset(0, 22).Value = month3ChangePercent
        rng2.Offset(0, 23).Value = month6ChangePercent
        rng2.Offset(0, 24).Value = ytdChangePercent
        rng2.Offset(0, 25).Value = year1ChangePercent
        rng2.Offset(0, 26).Value = year3ChangePercent
        rng2.Offset(0, 27).Value = year5ChangePercent
        rng2.Offset(0, 28).Value = beta
        rng2.Offset(0, 29).Value = marketcap
        rng2.Offset(0, 30).Value = sharesOutstanding
        rng2.Offset(0, 31).Value = float
        rng2.Offset(0, 32).Value = revenue
        rng2.Offset(0, 33).Value = revenuePerShare
        rng2.Offset(0, 34).Value = revenuePerEmployee
        rng2.Offset(0, 35).Value = EBITDA
        rng2.Offset(0, 36).Value = grossProfit
        rng2.Offset(0, 37).Value = profitMargin
        rng2.Offset(0, 38).Value = cash
        rng2.Offset(0, 39).Value = debt
        rng2.Offset(0, 40).Value = returnOnEquity
        rng2.Offset(0, 41).Value = returnOnAssets
        rng2.Offset(0, 42).Value = returnOnCapital
        rng2.Offset(0, 43).Value = peRatio
        rng2.Offset(0, 44).Value = peRatioLow
        rng2.Offset(0, 45).Value = peRatioHigh
        rng2.Offset(0, 46).Value = priceToSales
        rng2.Offset(0, 47).Value = priceToBook
        rng2.Offset(0, 48).Value = shortRatio
        rng2.Offset(0, 49).Value = grossProfit
        rng2.Offset(0, 50).Value = costOfRevenue
        rng2.Offset(0, 51).Value = opeartingRevenue
        rng2.Offset(0, 52).Value = totalRevenue
        rng2.Offset(0, 53).Value = opeartingIncome
        rng2.Offset(0, 54).Value = netIncome
        rng2.Offset(0, 55).Value = researchAndDevelopment
        rng2.Offset(0, 56).Value = opeartingExpenses
        rng2.Offset(0, 57).Value = currentAssets
        rng2.Offset(0, 58).Value = totalAssets
        rng2.Offset(0, 59).Value = totalLiabilities
        rng2.Offset(0, 60).Value = currentCash
        rng2.Offset(0, 61).Value = currentDebt
        rng2.Offset(0, 62).Value = totalCash
        rng2.Offset(0, 63).Value = totalDebt
        rng2.Offset(0, 64).Value = shareholderEquity
        rng2.Offset(0, 65).Value = cashChange
        rng2.Offset(0, 66).Value = cashFlow
        rng2.Offset(0, 67).Value = operatingGainsLosses
        rng2.Offset(0, 68).Value = amount
        rng2.Offset(0, 69).Value = dividendType
        rng2.Offset(0, 70).Value = dividendRate
        rng2.Offset(0, 71).Value = dividendYield
        rng2.Offset(0, 72).Value = exDate
        rng2.Offset(0, 73).Value = paymentDate
        rng2.Offset(0, 74).Value = declaredDate
        rng2.Offset(0, 75).Value = recordDate
        rng2.Offset(0, 76).Value = qualified
        rng2.Offset(1, 0).Select
        Set rng2 = ActiveCell
    Next
    j = j + 1
Wend
End If


End Sub
Sub extrastuff()

Dim tickerRange As Range
Dim tickerRangeLen As Integer
Dim rng As Range

Set tickerRange = Application.InputBox(prompt:="Select tickers", Type:=8)
x = tickerRange.Cells.Count
Set rng = tickerRange.Cells(1, 1)


If x >= 100 Then
y = x / 100
j = 1
For j = 1 To y
rng.Select

Dim tickers() As Variant
ReDim tickers(1 To 100) As Variant
Dim i As Integer

i = 1
For i = 1 To 100 Step 1
    tickers(i) = Selection.Value
    rng.Offset(1, 0).Select
    Set rng = ActiveCell
Next

Dim batch As String

batch = Join(tickers, ",")

Dim Dict As New Dictionary

Dict.CompareMode = CompareMethod.TextCompare
Dim companyName, exchange, sector, industry, CEO, issueType, dividendType As String
Dim latestPrice, openPrice, closePrice, low, high, change, changePercent, latestVolume, avgTotalVolume, week52Low, week52High, day50MovingAvg, day200MovingAvg, day5ChangePercent, month1ChangePercent, month3ChangePercent, month6ChangePercent, ytdChangePercent, year1ChangePercent, year3ChangePercent, year5ChangePercent, beta, marketcap, sharesOutstanding, float, revenue, revenuePerShare, revenuePerEmployee, EBITDA, grossProfit, profitMargin, cash, debt, returnOnEquity, returnOnAssets, returnOnCapital, peRatio, peRatioLow, peRatioHigh, priceToSales, priceToBook, shortRatio, costOfRevenue, opeartingRevenue, totalRevenue, opeartingIncome, netIncome, researchAndDevelopment, opeartingExpenses, currentAssets, totalAssets, totalLiabilities, currentCash, currentDebt, totalCash, totalDebt, shareholderEquity, cashChange, cashFlow, operatingGainsLosses, amount, dividendRate, dividendYield As Variant
Dim exDate, paymentDate, declaredDate, recordDate As Variant

Set rng = tickerRange.Cells(1, 1)

Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
MyRequest.Open "GET", "https://api.iextrading.com/1.0/stock/market/batch?symbols=" & batch & "&types=company,quote,stats,financials,earnings,dividends"
MyRequest.Send

Dim Json As Object

Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)

For i = 1 To 100 Step 1
    Dict("A") = rng.Value
    companyName = Json(Dict.Item("A"))("company")("companyName")
    exchange = Json(Dict.Item("A"))("company")("exchange")
    sector = Json(Dict.Item("A"))("company")("sector")
    industry = Json(Dict.Item("A"))("company")("industry")
    CEO = Json(Dict.Item("A"))("company")("CEO")
    issueType = Json(Dict.Item("A"))("company")("issueType")
    latestPrice = Json(Dict.Item("A"))("quote")("latestPrice")
    openPrice = Json(Dict.Item("A"))("quote")("open")
    closePrice = Json(Dict.Item("A"))("quote")("close")
    low = Json(Dict.Item("A"))("quote")("low")
    high = Json(Dict.Item("A"))("quote")("high")
    change = Json(Dict.Item("A"))("quote")("change")
    changePercent = Json(Dict.Item("A"))("quote")("changePercent")
    latestVolume = Json(Dict.Item("A"))("quote")("latestVolume")
    avgTotalVolume = Json(Dict.Item("A"))("quote")("avgTotalVolume")
    week52Low = Json(Dict.Item("A"))("quote")("week52Low")
    week52High = Json(Dict.Item("A"))("quote")("week52High")
    day50MovingAvg = Json(Dict.Item("A"))("stats")("day50MovingAvg")
    day200MovingAvg = Json(Dict.Item("A"))("stats")("day200MovingAvg")
    day5ChangePercent = Json(Dict.Item("A"))("stats")("day5ChangePercent")
    month1ChangePercent = Json(Dict.Item("A"))("stats")("month1ChangePercent")
    month3ChangePercent = Json(Dict.Item("A"))("stats")("month3ChangePercent")
    month6ChangePercent = Json(Dict.Item("A"))("stats")("month6ChangePercent")
    ytdChangePercent = Json(Dict.Item("A"))("stats")("ytdChangePercent")
    year1ChangePercent = Json(Dict.Item("A"))("stats")("year1ChangePercent")
    year3ChangePercent = Json(Dict.Item("A"))("stats")("year3ChangePercent")
    year5ChangePercent = Json(Dict.Item("A"))("stats")("year5ChangePercent")
    beta = Json(Dict.Item("A"))("stats")("beta")
    marketcap = Json(Dict.Item("A"))("stats")("marketcap")
    sharesOutstanding = Json(Dict.Item("A"))("stats")("sharesOutstanding")
    float = Json(Dict.Item("A"))("stats")("float")
    revenue = Json(Dict.Item("A"))("stats")("revenue")
    revenuePerShare = Json(Dict.Item("A"))("stats")("revenuePerShare")
    revenuePerEmployee = Json(Dict.Item("A"))("stats")("revenuePerEmployee")
    EBITDA = Json(Dict.Item("A"))("stats")("EBITDA")
    grossProfit = Json(Dict.Item("A"))("stats")("grossProfit")
    profitMargin = Json(Dict.Item("A"))("stats")("profitMargin")
    cash = Json(Dict.Item("A"))("stats")("cash")
    debt = Json(Dict.Item("A"))("stats")("debt")
    returnOnEquity = Json(Dict.Item("A"))("stats")("returnOnEquity")
    returnOnAssets = Json(Dict.Item("A"))("stats")("returnOnAssets")
    returnOnCapital = Json(Dict.Item("A"))("stats")("returnOnCapital")
    peRatio = Json(Dict.Item("A"))("quote")("peRatio")
    peRatioLow = Json(Dict.Item("A"))("stats")("peRatioLow")
    peRatioHigh = Json(Dict.Item("A"))("stats")("peRatioHigh")
    priceToSales = Json(Dict.Item("A"))("stats")("priceToSales")
    priceToBook = Json(Dict.Item("A"))("stats")("priceToBook")
    shortRatio = Json(Dict.Item("A"))("stats")("shortRatio")
    grossProfit = Json(Dict.Item("A"))("stats")("grossProfit")
    costOfRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("costOfRevenue")
    opeartingRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingRevenue")
    totalRevenue = Json(Dict.Item("A"))("financials")("financials")(1)("totalRevenue")
    opeartingIncome = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingIncome")
    netIncome = Json(Dict.Item("A"))("financials")("financials")(1)("netIncome")
    researchAndDevelopment = Json(Dict.Item("A"))("financials")("financials")(1)("researchAndDevelopment")
    opeartingExpenses = Json(Dict.Item("A"))("financials")("financials")(1)("opeartingExpenses")
    currentAssets = Json(Dict.Item("A"))("financials")("financials")(1)("currentAssets")
    totalAssets = Json(Dict.Item("A"))("financials")("financials")(1)("totalAssets")
    totalLiabilities = Json(Dict.Item("A"))("financials")("financials")(1)("totalLiabilities")
    currentCash = Json(Dict.Item("A"))("financials")("financials")(1)("currentCash")
    currentDebt = Json(Dict.Item("A"))("financials")("financials")(1)("currentDebt")
    totalCash = Json(Dict.Item("A"))("financials")("financials")(1)("totalCash")
    totalDebt = Json(Dict.Item("A"))("financials")("financials")(1)("totalDebt")
    shareholderEquity = Json(Dict.Item("A"))("financials")("financials")(1)("shareholderEquity")
    cashChange = Json(Dict.Item("A"))("financials")("financials")(1)("cashChange")
    cashFlow = Json(Dict.Item("A"))("financials")("financials")(1)("cashFlow")
    operatingGainsLosses = Json(Dict.Item("A"))("financials")("financials")(1)("operatingGainsLosses")

    rng.Offset(0, 1).Value = companyName
    rng.Offset(0, 2).Value = exchange
    rng.Offset(0, 3).Value = sector
    rng.Offset(0, 4).Value = industry
    rng.Offset(0, 5).Value = CEO
    rng.Offset(0, 6).Value = issueType
    rng.Offset(0, 7).Value = latestPrice
    rng.Offset(0, 8).Value = openPrice
    rng.Offset(0, 9).Value = closePrice
    rng.Offset(0, 10).Value = low
    rng.Offset(0, 11).Value = high
    rng.Offset(0, 12).Value = change
    rng.Offset(0, 13).Value = changePercent
    rng.Offset(0, 14).Value = latestVolume
    rng.Offset(0, 15).Value = avgTotalVolume
    rng.Offset(0, 16).Value = week52Low
    rng.Offset(0, 17).Value = week52High
    rng.Offset(0, 18).Value = day50MovingAvg
    rng.Offset(0, 19).Value = day200MovingAvg
    rng.Offset(0, 20).Value = day5ChangePercent
    rng.Offset(0, 21).Value = month1ChangePercent
    rng.Offset(0, 22).Value = month3ChangePercent
    rng.Offset(0, 23).Value = month6ChangePercent
    rng.Offset(0, 24).Value = ytdChangePercent
    rng.Offset(0, 25).Value = year1ChangePercent
    rng.Offset(0, 26).Value = year3ChangePercent
    rng.Offset(0, 27).Value = year5ChangePercent
    rng.Offset(0, 28).Value = beta
    rng.Offset(0, 29).Value = marketcap
    rng.Offset(0, 30).Value = sharesOutstanding
    rng.Offset(0, 31).Value = float
    rng.Offset(0, 32).Value = revenue
    rng.Offset(0, 33).Value = revenuePerShare
    rng.Offset(0, 34).Value = revenuePerEmployee
    rng.Offset(0, 35).Value = EBITDA
    rng.Offset(0, 36).Value = grossProfit
    rng.Offset(0, 37).Value = profitMargin
    rng.Offset(0, 38).Value = cash
    rng.Offset(0, 39).Value = debt
    rng.Offset(0, 40).Value = returnOnEquity
    rng.Offset(0, 41).Value = returnOnAssets
    rng.Offset(0, 42).Value = returnOnCapital
    rng.Offset(0, 43).Value = peRatio
    rng.Offset(0, 44).Value = peRatioLow
    rng.Offset(0, 45).Value = peRatioHigh
    rng.Offset(0, 46).Value = priceToSales
    rng.Offset(0, 47).Value = priceToBook
    rng.Offset(0, 48).Value = shortRatio
    rng.Offset(0, 49).Value = grossProfit
    rng.Offset(0, 50).Value = costOfRevenue
    rng.Offset(0, 51).Value = opeartingRevenue
    rng.Offset(0, 52).Value = totalRevenue
    rng.Offset(0, 53).Value = opeartingIncome
    rng.Offset(0, 54).Value = netIncome
    rng.Offset(0, 55).Value = researchAndDevelopment
    rng.Offset(0, 56).Value = opeartingExpenses
    rng.Offset(0, 57).Value = currentAssets
    rng.Offset(0, 58).Value = totalAssets
    rng.Offset(0, 59).Value = totalLiabilities
    rng.Offset(0, 60).Value = currentCash
    rng.Offset(0, 61).Value = currentDebt
    rng.Offset(0, 62).Value = totalCash
    rng.Offset(0, 63).Value = totalDebt
    rng.Offset(0, 64).Value = shareholderEquity
    rng.Offset(0, 65).Value = cashChange
    rng.Offset(0, 66).Value = cashFlow
    rng.Offset(0, 67).Value = operatingGainsLosses
    rng.Offset(1, 0).Select
    Set rng = ActiveCell
Next
Next
End If
End Sub
