Attribute VB_Name = "GetPrice"
Option Explicit
Dim priceLoopCount As Integer

'Private Function eval(str As String)
'    'str = Replace(str, """", "'")
'    'str = Replace(str, ";", "")
'    eval = Application.evaluate(str)
'End Function

Private Function ArrangeDates(ByVal startDate As Date, ByVal endDate As Date) As Variant()
    
    If startDate <> 0 Then
        If WorksheetFunction.Weekday(startDate, 2) > 5 Then
            startDate = DateAdd("d", -(WorksheetFunction.Weekday(startDate, 2) - 5), startDate)
        Else
            
        End If
        If startDate > Date Then startDate = Date
    Else
        startDate = Date
    End If
    
    If endDate <> 0 Then
        If WorksheetFunction.Weekday(endDate, 2) > 5 Then
            endDate = DateAdd("d", (8 - WorksheetFunction.Weekday(endDate, 2)), endDate)
        Else
            
        End If
        If endDate > Date Then endDate = Date
    Else
        'No end date only one result will be shown
        endDate = 0
    End If
    If startDate = endDate Then endDate = 0
    
    Dim result(0 To 1)
    result(0) = startDate
    result(1) = endDate
    ArrangeDates = result
    
End Function

Public Function PriceThat(name As String, Optional theType As String, Optional startDate As Date, Optional endDate As Date) As Variant
   
    Application.EnableEvents = False
    '
    'Weekend Protector.
    Dim dates As Variant
    dates = ArrangeDates(startDate, endDate)
    Dim startDateVal As Date
    Dim endDateVal As Date
    startDateVal = dates(0)
    endDateVal = dates(1)
    
    ' PriceThat = InvestingPrice(name, daterangeval)
    'Exit Function
    If UCase(theType) = "TEFAS" Then
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = FundCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = FundPrice(name, startDateVal)
        Else
            PriceThat = FundPrice(name, startDateVal, endDateVal)
        End If
    ElseIf UCase(theType) = "DÖVÝZ" Then
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = CommodityCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = CommodityPrice(name, startDateVal)
        Else
            PriceThat = CommodityPrice(name, startDateVal, endDateVal)
        End If
    ElseIf UCase(theType) = "HÝSSE" Then
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = StockCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = StockPrice(name, startDateVal)
        Else
            PriceThat = StockPrice(name, startDateVal, endDateVal)
        End If
        If TypeName(PriceThat) <> "Variant()" Then If PriceThat = 0 Then PriceThat = InvestingPrice(name, startDateVal, endDateVal)
        If TypeName(PriceThat) = "Variant()" Then If IsEmpty(PriceThat) Then PriceThat = InferMarketPrices(name, startDateVal, endDateVal)
    ElseIf left(name, 2) = "W_" Or name = "REPO" Then
        PriceThat = 0
    ElseIf Len(name) > 3 Then
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = StockCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = StockPrice(name, startDateVal)
        Else
            PriceThat = StockPrice(name, startDateVal, endDateVal)
        End If
        If TypeName(PriceThat) <> "Variant()" Then If PriceThat = 0 Then PriceThat = InvestingPrice(name, startDateVal, endDateVal)
        If TypeName(PriceThat) = "Variant()" Then If IsEmpty(PriceThat) Then PriceThat = InferMarketPrices(name, startDateVal, endDateVal)
    ElseIf UCase(name) = "USD" Or UCase(name) = "EUR" Or UCase(name) = "GBP" Then
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = CommodityCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = CommodityPrice(name, startDateVal)
        Else
            PriceThat = CommodityPrice(name, startDateVal, endDateVal)
        End If
    Else
        If startDateVal = Date And endDateVal = 0 Then
            PriceThat = FundCurrentPrice(name)
        ElseIf endDateVal = 0 Then
            PriceThat = FundPrice(name, startDateVal)
        Else
            PriceThat = FundPrice(name, startDateVal, endDateVal)
        End If
    End If
    
    Application.EnableEvents = True
    'Application.Calculation = xlCalculationAutomatic
    
    ' it needs to be worked.... if there is a fund like that in the time or we cun run the AccSummary first.
    'If PriceThat = 0 And priceLoopCount < 2 Then
    '    priceLoopCount = priceLoopCount + 1
    '    PriceThat = PriceThat(name, dateRange - 1)
    'Else
    'End If
    
    'priceLoopCount = 0

End Function
Private Function InferMarketPrices(name As String, startDateVal As Date, Optional endDateVal As Date) As Variant
    
    'find the last price
    Dim rownum As Long
    rownum = 2 'starting row
    Dim ddate  As Date
    Dim entityName As String
    Dim entityPrice As Double
    Dim lowestDate As Date
    lowestDate = DateAdd("m", 5, Date)
    With TEFAS_PRICES
        Do While .Cells(rownum, 1) <> ""
            ddate = .Cells(rownum, 1)
            entityName = .Cells(rownum, 3)
            If entityName = name And ddate < lowestDate Then
                entityPrice = .Cells(rownum, 5)
                lowestDate = ddate
            Else
            End If
            rownum = rownum + 1
'            If rownum = 37659 Then Stop
'            Debug.Print rownum
        Loop
    End With
    'create array
    Dim result() As Variant
    Dim i As Integer
    ReDim result(endDateVal - startDateVal, 0 To 1)
    For i = LBound(result) To UBound(result)
        result(i, 0) = DateAdd("d", i, startDateVal)
        result(i, 1) = entityPrice
    Next i
    
    InferMarketPrices = result
    
End Function
Private Function IsWeekend(InputDate As Date) As Boolean
    Select Case Weekday(InputDate)
    Case vbSaturday, vbSunday
        IsWeekend = True
    Case Else
        IsWeekend = False
    End Select
End Function

Private Function CommodityPrice(fromCurrSym As String, startDateVal As Date, Optional endDateVal As Date)

    fromCurrSym = left(UCase(fromCurrSym), 3)
    Dim theLink As String
    Dim HttpReq As Variant
    Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    Dim httpResponse As String
    If endDateVal = 0 Then
        'shows the currency from https://www.tcmb.gov.tr/kurlar/202107/07072021.xml?_=162
        theLink = "https://www.tcmb.gov.tr/kurlar/"
        theLink = theLink & Year(startDateVal)
        If Month(startDateVal) < 10 Then
            theLink = theLink & "0" & Month(startDateVal) & "/"
        Else
            theLink = theLink & Month(startDateVal) & "/"
        End If
        If Day(startDateVal) < 10 Then
            theLink = theLink & "0" & Day(startDateVal)
        Else
            theLink = theLink & Day(startDateVal)
        End If
        If Month(startDateVal) < 10 Then
            theLink = theLink & "0" & Month(startDateVal) & Year(startDateVal) & ".xml?_="
        Else
            theLink = theLink & Month(startDateVal) & Year(startDateVal) & ".xml?_="
        End If
        theLink = theLink & WorksheetFunction.RandBetween(1000000000000#, 1999999999999#)
        'Dim httpreq As MSXML2.ServerXMLHTTP
        HttpReq.Open "GET", theLink, False
        HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        HttpReq.send
        httpResponse = HttpReq.responseText
        Dim target As String
        If InStr(httpResponse, "Kod=""" & fromCurrSym) = 0 Then
            CommodityPrice = 0 'if currency cannot be found
        Else
            target = right(httpResponse, Len(httpResponse) - InStr(httpResponse, "Kod=""" & fromCurrSym))
            target = right(target, Len(target) - InStr(target, "<ForexBuying>") - 12)
            target = left(target, InStr(target, "</ForexBuying>") - 1)
            target = Replace(target, ".", ",")
            CommodityPrice = CDbl(target)
            Set HttpReq = Nothing
            Exit Function
        End If
    Else
        'if enddate is not missing
        theLink = "https://www.bloomberght.com/piyasa/refdata/dolar"
        HttpReq.Open "GET", theLink, False
        HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        HttpReq.send
        httpResponse = HttpReq.responseText
        Dim vson As Variant
        Dim sson As String
        JSON.Parse HttpReq.responseText, vson, sson
        If sson = "Error" Then
            CommodityPrice = 0
        Else
            Dim endIndex As Long
            Dim startIndex As Long
            vson = vson("SeriesData")
            For startIndex = LBound(vson) To UBound(vson)
                If CDate(((CDbl(vson(startIndex)(0)) + 10800000) / 1000 / 60 / 60 / 24) + 25569) >= startDateVal Then Exit For
            Next startIndex
            startIndex = startIndex - 0
            For endIndex = UBound(vson) To LBound(vson) Step -1
                If CDate(((CDbl(vson(endIndex)(0)) + 10800000) / 1000 / 60 / 60 / 24) + 25569) <= endDateVal Then Exit For
            Next endIndex
            endIndex = endIndex - 0
            Dim result() As Variant
            ReDim result(0 To endIndex - startIndex, 0 To 1)
            Dim i As Integer
            Dim j As Integer
            For i = startIndex To endIndex
                result(j, 0) = CDate(((CDbl(vson(i)(0)) + 10800000) / 1000 / 60 / 60 / 24) + 25569)
                result(j, 1) = CDbl(vson(i)(1))
                j = j + 1
            Next i
            CommodityPrice = result
            Exit Function
        End If
    End If

End Function
Private Function CommodityCurrentPrice(fromCurrSym As String)

    Dim theLink As String
    theLink = "https://www.bloomberght.com/doviz/"
    fromCurrSym = UCase(fromCurrSym)
    Select Case fromCurrSym
    Case "USD"
        theLink = theLink & "dolar"
    Case "EUR"
        theLink = theLink & "euro"
    Case "GBP"
        theLink = theLink & "ingiliz-sterlini"
    Case "JPY"
        theLink = theLink & "japon-yeni"
    Case Else
        CommodityCurrentPrice = "BÝLÝNMEYEN ANLIK SEMBOL"
        Exit Function
    End Select
    
    Dim HttpReq As Variant
    Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    'added cb, a random variable for cache persistency issues. i don't want the value to be cached.
    HttpReq.Open "GET", theLink & "?cb=" & Timer() * 100, False
    HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    HttpReq.setRequestHeader "Cache-Control", "no-cache"
    HttpReq.setRequestHeader "pragma", "no-cache"
    HttpReq.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    HttpReq.send
    
    Dim httpResponse As String
    httpResponse = HttpReq.responseText
    Set HttpReq = Nothing
    Dim htmlDoc As New HTMLDocument
    htmlDoc.body.innerHTML = httpResponse
    
    Dim target As Object
    Set target = htmlDoc.getElementsByClassName("widget-interest-detail type1")
    Set target = target(0).children(0).children(0).children(1)
    CommodityCurrentPrice = CDbl(Mid(target.innerText, InStr(1, target.innerText, ")") + 2, 7))

End Function

Private Function FundCurrentPrice(fundCode As String) 'rng As Range)
    Dim HttpReq As New MSXML2.XMLHTTP60
    Dim package As String
    Dim theLink As String
    On Error GoTo 0
    package = fundCode
    theLink = "https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod="
    With HttpReq
        .Open "GET", theLink & package, False
        .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
        .send
    End With
    
    Dim htmlDoc As New MSHTML.HTMLDocument
    htmlDoc.body.innerHTML = HttpReq.responseText
    
    Dim strStart As Long
    Dim strLen As Long
    Dim q As String
    strStart = InStr(1, HttpReq.responseText, "chartMainContent_FonFiyatGrafik = new Highcharts.Chart(")
    strLen = InStr(strStart, HttpReq.responseText, "});") - strStart
    q = Mid(HttpReq.responseText, strStart + 55, strLen - 54)
    q = Replace(q, " ", "")
    q = Replace(q, vbNewLine, "")
    q = "[" & q & "]"
    'q = Left(q, Len(q) - 3)
    'q = q & "}]"
    q = Replace(q, "{""formatter"":function(event){vartmp='<b>'+this.series.name+((typeof(this.point.name)!='undefined')?'->'+this.point.name:'')+'</b><br/>'+this.x+':'+this.y;if(typeof(tmp)=='function'){returntmp(this);}else{returntmp;}}}", "0")
    Dim vson As Variant
    Dim sson As String
    JSON.Parse q, vson, sson
    If sson = "Error" Then
        FundCurrentPrice = 0
    Else
        vson = vson(0)("series")
        If UBound(vson) = -1 Then
            FundCurrentPrice = 0
        Else
            vson = vson(0)("data")
            vson = CDbl(vson(UBound(vson)))
            FundCurrentPrice = CDbl(vson)
        End If
    End If
End Function

Private Function FundPrice(fundCode As String, startDateVal As Date, Optional endDateVal As Date)
    Dim HttpReq As New MSXML2.XMLHTTP60
    Dim package As String
    Dim theLink As String
    Dim startDate As String
    Dim endDate As String
    Dim startDateStamp As Long
    Dim endDateStamp As Long
    Dim finalResult() As Variant
    ReDim finalResult(0, 1)
    Dim tempResult() As Variant
    ReDim tempResult(0, 1)
        
    Dim periodDiff As Integer
    startDateStamp = CLng(Year(startDateVal) & IIf(Month(startDateVal) < 10, "0", "") & Month(startDateVal))
    endDateStamp = CLng(Year(endDateVal) & IIf(Month(endDateVal) < 10, "0", "") & Month(endDateVal))
    If endDateVal = 0 Then
        periodDiff = 0
    Else
        periodDiff = ((Year(endDateVal) - Year(startDateVal)) * 4) + Int((Month(endDateVal) - Month(startDateVal) - 1) / 3)
        If periodDiff < 0 Then periodDiff = 0
    End If
    
    Dim j As Integer
    For j = 0 To periodDiff
        startDate = CStr(DateAdd("m", 3 * j, startDateVal))
        If endDateVal = 0 Then
            endDate = startDate
        Else
            If DateAdd("m", 3 * (j + 1), startDateVal) >= endDateVal Then
                endDate = CStr(endDateVal)
            Else
                endDate = CStr(DateAdd("m", 3 * (j + 1), startDateVal))
            End If
        End If
        package = "fontip=YAT&sfontur=&fonkod=" + fundCode + "&fongrup=&bastarih=" + startDate + "&bittarih=" + endDate + "&fonturkod=&fonunvantip="
        theLink = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
        With HttpReq
            .Open "POST", theLink, False
            .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
            .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
            .send package
        End With
        
        Dim vson As Variant
        Dim sson As String
        JSON.Parse HttpReq.responseText, vson, sson
        If sson = "Error" Then
            FundPrice = 0
        Else
            vson = vson("data")
            If UBound(vson) = -1 Then
                FundPrice = 0
            Else
                If endDateVal = 0 Then
                    vson = vson(0)("FIYAT")
                    FundPrice = CDbl(vson)
                    Exit Function
                Else
                    Dim result() As Variant
                    ReDim result(LBound(vson) To UBound(vson), 0 To 1)
                    Dim i As Integer
                    For i = UBound(vson) To LBound(vson) Step -1
                        result(UBound(vson) - i, 0) = CDate((CDbl(vson(i)("TARIH")) / 1000 / 60 / 60 / 24) + 25569)
                        result(UBound(vson) - i, 1) = CDbl(vson(i)("FIYAT"))
                    Next i
                    FundPrice = result
                End If
            End If
        End If
        'Result aggregator
        tempResult = finalResult
        ReDim finalResult(LBound(finalResult) To UBound(finalResult) + UBound(result), 0 To 1)
        For i = LBound(tempResult) To UBound(tempResult)
            finalResult(i, 0) = tempResult(i, 0)
            finalResult(i, 1) = tempResult(i, 1)
        Next i
        For i = LBound(result) To UBound(result)
            finalResult(i + UBound(tempResult), 0) = result(i, 0)
            finalResult(i + UBound(tempResult), 1) = result(i, 1)
        Next i
    Next j
    
    FundPrice = finalResult

End Function

Private Function StockCurrentPrice(stockCode As String)
Dim HttpReq As New MSXML2.XMLHTTP60
Dim package As String
Dim theLink As String

package = stockCode & ".E.BIST"
theLink = "https://www.isyatirim.com.tr/_layouts/15/Isyatirim.Website/Common/Data.aspx/OneEndeks?endeks="
With HttpReq
    .Open "GET", theLink & package, False
    .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    .send
End With

Dim vson As Variant
Dim sson As String
JSON.Parse HttpReq.responseText, vson, sson
If sson = "Error" Then
    StockCurrentPrice = 0
Else
    'vson = vson(0)
    'If UBound(vson) = -1 Then
    '    StockCurrentPrice = 0
    'Else
        vson = vson(0)("last")
        StockCurrentPrice = CDbl(vson)
    'End If
End If

End Function

Private Function StockPrice(stockCode As String, startDateVal As Date, Optional endDateVal As Date)
Dim HttpReq As New MSXML2.XMLHTTP60
Dim theLink As String
Dim startDate As String
Dim endDate As String

startDate = Replace(startDateVal, ".", "-")
If Day(startDateVal) < 10 Then startDate = "0" & startDate
If endDateVal = 0 Then
    endDate = startDate
Else
    endDate = Replace(endDateVal, ".", "-")
    If Day(endDateVal) < 10 Then endDate = "0" & endDate
End If

theLink = "https://www.isyatirim.com.tr/_layouts/15/Isyatirim.Website/Common/Data.aspx/HisseTekil?hisse=" + stockCode + "&startdate=" + startDate + "&enddate=" + endDate
With HttpReq
    .Open "GET", theLink, False
    .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    .send
End With

Dim vson As Variant
Dim sson As String
JSON.Parse HttpReq.responseText, vson, sson

If sson = "Error" Then
    StockPrice = 0
Else
    vson = vson("value")
    If UBound(vson) = -1 Then
        StockPrice = 0
    Else
        Dim result() As Variant
        Dim i As Integer
        If UBound(vson) = 0 And endDateVal = 0 Then
            vson = vson(0)("HGDG_KAPANIS") '("END_DEGER") 'HGDG_TARIH
            StockPrice = CDbl(vson)
        ElseIf UBound(vson) = 0 And endDateVal <> 0 Then
            ReDim result(0, 0 To 1)
            For i = LBound(result) To UBound(result)
                result(i, 0) = CDate(DateAdd("d", i, startDateVal))
                result(i, 1) = CDbl(vson(0)("HGDG_KAPANIS"))
            Next i
            StockPrice = result
        Else
            ReDim result(LBound(vson) To UBound(vson), 0 To 1)
            For i = LBound(vson) To UBound(vson)
                result(i, 0) = CDate(vson(i)("HGDG_TARIH"))
                result(i, 1) = CDbl(vson(i)("HGDG_KAPANIS"))
            Next i
            StockPrice = result
        End If
    End If
End If

End Function

Private Function InvestingPrice(stockCode As String, startDateVal As Date, Optional endDateVal As Date) As Variant
    Dim edge As New CDPBrowser
    Dim theLink As String
    'Start and hide
    edge.start name:="edge", reAttach:=False
    edge.hide
    theLink = "https://api.investing.com/api/search/v2/search?q=" & stockCode
    'Perform automation in the background
    edge.navigate theLink, isInteractive
    
    Dim inside As Variant
    Set inside = edge.getElementsByXPath("/html/body/div[3]/div[2]/div[2]")
    Dim vson As Variant
    Dim state As String
    Dim quoteId As String
    Call JSON.Parse(inside(1).innerText, vson, state)
    On Error GoTo errquit
    If state = "Error" Then
        Exit Function
    Else
        quoteId = vson("quotes")(0)("id")
    End If
    On Error GoTo 0
    theLink = "https://api.investing.com/api/financialdata/" & quoteId & "/historical/chart/?period=MAX&interval=P1D&pointscount=160"
    edge.navigate theLink, isInteractive
    
    inside = edge.jsEval("document.documentElement.outerHTML")
    Dim str As String
    str = Mid(inside, InStr(1, inside, "{""data"":"))
    str = left(str, InStr(1, str, """events"":null}") + 13)
    Call JSON.Parse(str, vson, state)
    If state = "Error" Then
        InvestingPrice = 0
    Else
        Dim endIndex As Long
        Dim startIndex As Long
        vson = vson("data")
        For startIndex = LBound(vson) To UBound(vson)
            If CDate(((CDbl(vson(startIndex)(0)) + 0) / 1000 / 60 / 60 / 24) + 25569) >= startDateVal Then Exit For
        Next startIndex
        startIndex = startIndex - 0
        If endDateVal = 0 Then
            endIndex = startIndex
        Else
            For endIndex = UBound(vson) To LBound(vson) Step -1
                If CDate(((CDbl(vson(endIndex)(0)) + 0) / 1000 / 60 / 60 / 24) + 25569) <= endDateVal Then Exit For
            Next endIndex
            endIndex = endIndex - 0
        End If
        If endDateVal = 0 Then
            InvestingPrice = CDbl(vson(startIndex)(4))
        Else
            Dim result() As Variant
            On Error GoTo errquit
            ReDim result(0 To endIndex - startIndex, 0 To 1)
            Dim i As Integer
            Dim j As Integer
            For i = startIndex To endIndex
                result(j, 0) = CDate(((CDbl(vson(i)(0)) + 0) / 1000 / 60 / 60 / 24) + 25569)
                result(j, 1) = CDbl(vson(i)(4))
                j = j + 1
            Next i
            InvestingPrice = result
        End If
    End If
errquit:
    
    edge.quit
    
End Function

thirty", "marginoferror" : "pointfive", "sampleratepertenthousand" : 0 }, "config78" : { "eventfrequency" : "seventeentolessthanthirty", "marginoferror" : "one", "sampleratepertenthousand" : 0 }, "config79" : { "eventfrequency" : "thirtytolessthanfiftyfive", "marginoferror" : "pointzeroone", "sampleratepertenthousand" : 100 }, "config80" : { "eventfrequency" : "thirtytolessthanfiftyfive", "marginoferror" : "pointzerotwo", "sampleratepertenthousand" : 10 }, "config81" : { "eventfrequency" : "thirtytolessthan