Sub stockScrapeAlpha()

'define ticker range

Dim tickerRange As Range
Set tickerRange = Application.InputBox(prompt:="Select tickers", Type:=8)

'count the number of cells in tickerRange and store that in an int

Dim tickerRangeLen As Integer
tickerRangeLen = tickerRange.Cells.Count
MsgBox tickerRangeLen

'define the top two ranges that both identify the first cell in the column

Dim Rng1 As Range
Dim Rng2 As Range
Set Rng1 = tickerRange.Cells(1, 1)
Set Rng2 = tickerRange.Cells(1, 1)

'define a batch variable

Dim batch As String

'define an array for the tickers

Dim tickers() As Variant

'create a JSON object
Dim Json As Object

'create a dicitonary
Dim Dict As New Dictionary
Dict.CompareMode = CompareMethod.TextCompare

'The max number of tickers per request is 100
'SOOOO we need to define some extra stuff if you happen to be fetching more than 100 tickers

Dim qtyHundredBatches As Integer
Dim remainder As Integer
Dim i As Integer
Dim j As Integer

qtyHundredBatches = tickerRangeLen / 100
remainder = tickerRangeLen Mod 100

If tickerRangeLen >= 100 Then
j = 1
While j <= qtyHundredBatches
    
    ReDim tickers(1 To 100) As Variant
    
    'push a hundred values into the array
    For i = 1 To 100 Step 1
        Rng1.Select
        tickers(i) = Selection.Value
        Rng1.Offset(1, 0).Select
        Set Rng1 = ActiveCell
    Next
    
    'join those hundred into a single string string
    batch = Join(tickers, ",")
    
    'fetch the url
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", "https://api.iextrading.com/1.0/stock/market/batch?symbols=" & batch & "&types=company,quote,stats,financials,earnings,dividends"
    MyRequest.Send
    
    'create a JSON object
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    
    'paste the JSON values into spreasheet
    
    For i = 1 To 100 Step 1
        Dict("A") = Rng2.Value
        Call iexTradingJSON(Dict, Rng1, Rng2, Json)
        Set Rng2 = ActiveCell
    Next
    j = j + 1
Wend
    'redefine tickers
    ReDim tickers(1 To remainder) As Variant

    'push a hundred values into an array
    For i = 1 To remainder Step 1
        Rng1.Select
        tickers(i) = Selection.Value
        Rng1.Offset(1, 0).Select
        Set Rng1 = ActiveCell
    Next
    
    'join those hundred into a single string string
    batch = Join(tickers, ",")
    
    'fetch the url
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", "https://api.iextrading.com/1.0/stock/market/batch?symbols=" & batch & "&types=company,quote,stats,financials,earnings,dividends"
    MyRequest.Send
    
    'Set JSON
    
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    
    'paste the JSON values into spreasheet
    
    For i = 1 To remainder Step 1
        Dict("A") = Rng2.Value
        Call iexTradingJSON(Dict, Rng1, Rng2, Json)
        Set Rng2 = ActiveCell
    Next

End If
If tickerRangeLen < 100 Then
    
    'redefine tickers
    ReDim tickers(1 To tickerRangeLen) As Variant

    'push values into an array
    For i = 1 To tickerRangeLen Step 1
        Rng1.Select
        tickers(i) = Selection.Value
        Rng1.Offset(1, 0).Select
        Set Rng1 = ActiveCell
    Next
    
    'join those hundred into a single string string
    batch = Join(tickers, ",")
    
    'fetch the url
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", "https://api.iextrading.com/1.0/stock/market/batch?symbols=" & batch & "&types=company,quote,stats,financials,earnings,dividends"
    MyRequest.Send
    
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    
    'paste the JSON values into spreasheet
    
    For i = 1 To tickerRangeLen Step 1
        Dict("A") = Rng2.Value
        Call iexTradingJSON(Dict, Rng1, Rng2, Json)
        Set Rng2 = ActiveCell
    Next
End If

End Sub


Public Function iexTradingJSON(Dict As Dictionary, Rng1 As Range, Rng2 As Range, Json As Object)

    Dim companyName, exchange, sector, industry, CEO, issueType, dividendType As Variant
    Dim latestPrice, openPrice, closePrice, low, high, change, changePercent, latestVolume, avgTotalVolume, week52Low, week52High, day50MovingAvg, day200MovingAvg, day5ChangePercent, month1ChangePercent, month3ChangePercent, month6ChangePercent, ytdChangePercent, year1ChangePercent, year3ChangePercent, year5ChangePercent, beta, marketcap, sharesOutstanding, float, revenue, revenuePerShare, revenuePerEmployee, EBITDA, grossProfit, profitMargin, cash, debt, returnOnEquity, returnOnAssets, returnOnCapital, peRatio, peRatioLow, peRatioHigh, priceToSales, priceToBook, shortRatio, costOfRevenue, opeartingRevenue, totalRevenue, opeartingIncome, netIncome, researchAndDevelopment, opeartingExpenses, currentAssets, totalAssets, totalLiabilities, currentCash, currentDebt, totalCash, totalDebt, shareholderEquity, cashChange, cashFlow, operatingGainsLosses, amount, dividendRate, dividendYield As Variant
    Dim exDate, paymentDate, declaredDate, recordDate As Variant
        
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
        Rng2.Offset(0, 1).Value = companyName
        Rng2.Offset(0, 2).Value = exchange
        Rng2.Offset(0, 3).Value = sector
        Rng2.Offset(0, 4).Value = industry
        Rng2.Offset(0, 5).Value = CEO
        Rng2.Offset(0, 6).Value = issueType
        Rng2.Offset(0, 7).Value = latestPrice
        Rng2.Offset(0, 8).Value = openPrice
        Rng2.Offset(0, 9).Value = closePrice
        Rng2.Offset(0, 10).Value = low
        Rng2.Offset(0, 11).Value = high
        Rng2.Offset(0, 12).Value = change
        Rng2.Offset(0, 13).Value = changePercent
        Rng2.Offset(0, 14).Value = latestVolume
        Rng2.Offset(0, 15).Value = avgTotalVolume
        Rng2.Offset(0, 16).Value = week52Low
        Rng2.Offset(0, 17).Value = week52High
        Rng2.Offset(0, 18).Value = day50MovingAvg
        Rng2.Offset(0, 19).Value = day200MovingAvg
        Rng2.Offset(0, 20).Value = day5ChangePercent
        Rng2.Offset(0, 21).Value = month1ChangePercent
        Rng2.Offset(0, 22).Value = month3ChangePercent
        Rng2.Offset(0, 23).Value = month6ChangePercent
        Rng2.Offset(0, 24).Value = ytdChangePercent
        Rng2.Offset(0, 25).Value = year1ChangePercent
        Rng2.Offset(0, 26).Value = year3ChangePercent
        Rng2.Offset(0, 27).Value = year5ChangePercent
        Rng2.Offset(0, 28).Value = beta
        Rng2.Offset(0, 29).Value = marketcap
        Rng2.Offset(0, 30).Value = sharesOutstanding
        Rng2.Offset(0, 31).Value = float
        Rng2.Offset(0, 32).Value = revenue
        Rng2.Offset(0, 33).Value = revenuePerShare
        Rng2.Offset(0, 34).Value = revenuePerEmployee
        Rng2.Offset(0, 35).Value = EBITDA
        Rng2.Offset(0, 36).Value = grossProfit
        Rng2.Offset(0, 37).Value = profitMargin
        Rng2.Offset(0, 38).Value = cash
        Rng2.Offset(0, 39).Value = debt
        Rng2.Offset(0, 40).Value = returnOnEquity
        Rng2.Offset(0, 41).Value = returnOnAssets
        Rng2.Offset(0, 42).Value = returnOnCapital
        Rng2.Offset(0, 43).Value = peRatio
        Rng2.Offset(0, 44).Value = peRatioLow
        Rng2.Offset(0, 45).Value = peRatioHigh
        Rng2.Offset(0, 46).Value = priceToSales
        Rng2.Offset(0, 47).Value = priceToBook
        Rng2.Offset(0, 48).Value = shortRatio
        Rng2.Offset(0, 49).Value = grossProfit
        Rng2.Offset(0, 50).Value = costOfRevenue
        Rng2.Offset(0, 51).Value = opeartingRevenue
        Rng2.Offset(0, 52).Value = totalRevenue
        Rng2.Offset(0, 53).Value = opeartingIncome
        Rng2.Offset(0, 54).Value = netIncome
        Rng2.Offset(0, 55).Value = researchAndDevelopment
        Rng2.Offset(0, 56).Value = opeartingExpenses
        Rng2.Offset(0, 57).Value = currentAssets
        Rng2.Offset(0, 58).Value = totalAssets
        Rng2.Offset(0, 59).Value = totalLiabilities
        Rng2.Offset(0, 60).Value = currentCash
        Rng2.Offset(0, 61).Value = currentDebt
        Rng2.Offset(0, 62).Value = totalCash
        Rng2.Offset(0, 63).Value = totalDebt
        Rng2.Offset(0, 64).Value = shareholderEquity
        Rng2.Offset(0, 65).Value = cashChange
        Rng2.Offset(0, 66).Value = cashFlow
        Rng2.Offset(0, 67).Value = operatingGainsLosses
        Rng2.Offset(0, 68).Value = amount
        Rng2.Offset(0, 69).Value = dividendType
        Rng2.Offset(0, 70).Value = dividendRate
        Rng2.Offset(0, 71).Value = dividendYield
        Rng2.Offset(0, 72).Value = exDate
        Rng2.Offset(0, 73).Value = paymentDate
        Rng2.Offset(0, 74).Value = declaredDate
        Rng2.Offset(0, 75).Value = recordDate
        Rng2.Offset(0, 76).Value = qualified
        Rng2.Offset(1, 0).Select

End Function
