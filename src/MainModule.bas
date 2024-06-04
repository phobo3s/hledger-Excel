Attribute VB_Name = "MainModule"
Option Explicit
Dim commDict As Object

'main column Ledger headers
Dim dateCol As Integer
Dim transCodeCol As Integer
Dim descriptionCol As Integer
Dim notesCol As Integer
Dim currencyCol As Integer
Dim operationCol As Integer
Dim tagCol As Integer
Dim accNameCol As Integer
Dim amountCol As Integer
Dim rateCol As Integer
Dim reconCol As Integer
    
Sub CreateAllFilesAKATornado()
    
    'aggregate all sub account pages
    Call AggregateAccounts(PopulateAccountsSheets)
    'main hledger file creator module
    Call ExportHledgerFile
    'MAIN_LEDGER.Activate
    'Call ExportAsCSV
    'export commodity prices
    ''TEFAS_PRICES.activate
    ''Call ExportAsCSV
    MAIN_LEDGER.activate
End Sub

Private Sub PopulateColumnHeaderIndexes(pageName As String)
    Dim sh As Worksheet
    Dim shName As String
    
    dateCol = 0
    transCodeCol = 0
    descriptionCol = 0
    notesCol = 0
    currencyCol = 0
    operationCol = 0
    tagCol = 0
    accNameCol = 0
    amountCol = 0
    rateCol = 0
    reconCol = 0
        
    For Each sh In ThisWorkbook.Worksheets
        If sh.CodeName = pageName Then shName = sh.name: Exit For
    Next sh
    On Error Resume Next
    With ThisWorkbook.Worksheets(shName)
        dateCol = .Cells(1, 1).EntireRow.Find("Date").Column
        transCodeCol = .Cells(1, 1).EntireRow.Find("Transaction Code").Column
        descriptionCol = .Cells(1, 1).EntireRow.Find("Payee|Note").Column
        notesCol = .Cells(1, 1).EntireRow.Find("Notes").Column
        currencyCol = .Cells(1, 1).EntireRow.Find("Commodity/Currency").Column
        operationCol = .Cells(1, 1).EntireRow.Find("Operation").Column
        tagCol = .Cells(1, 1).EntireRow.Find("Tag/Note").Column
        accNameCol = .Cells(1, 1).EntireRow.Find("Full Account Name").Column
        amountCol = .Cells(1, 1).EntireRow.Find("Amount").Column
        rateCol = .Cells(1, 1).EntireRow.Find("Rate/Price").Column
        reconCol = .Cells(1, 1).EntireRow.Find("Reconciliation").Column
    End With
    On Error GoTo 0
End Sub
Private Function PopulateAccountsSheets() As Variant

    Dim result(1 To 6)
    result(1) = "TEBKrediKartý"
    result(2) = "TEBKrediKartý-2911"
    result(3) = "SheKrediKartý"
    result(4) = "TEBBanka"
    result(5) = "Nakit"
    result(6) = "IS-Yatýrým"
    PopulateAccountsSheets = result
    
End Function
Private Sub AggregateAccounts(accountsArr As Variant)

    Dim lastRowNum As Long
    lastRowNum = MAIN_LEDGER.Range("H2").End(xlDown).Row
    MAIN_LEDGER.Range("A2").Resize(lastRowNum - 1, 12).value = ""
    MAIN_LEDGER.Range("A2").Resize(lastRowNum - 1, 12).Interior.ColorIndex = -4142
    Dim sh As Worksheet
    Dim i As Integer
    For i = LBound(accountsArr) To UBound(accountsArr)
        Set sh = Application.ActiveWorkbook.Worksheets(accountsArr(i))
        lastRowNum = sh.Cells(sh.Cells.Rows.Count, sh.Range("H2").Column).End(xlUp).Row
        If lastRowNum <> 1 Then
            sh.Cells(2, 1).Resize(lastRowNum - 1, 12).Copy _
            MAIN_LEDGER.Cells(MAIN_LEDGER.Cells(MAIN_LEDGER.Cells.Rows.Count, MAIN_LEDGER.Range("H2").Column).End(xlUp).Row + 1, 1)
        Else
        End If
    Next i
    
End Sub

Public Sub GetNewPriceData()
    
    ' > Address column headers
    Call PopulateColumnHeaderIndexes("TEFAS_PRICES")
    ' < Address column headers
    
    Dim commodityRange As Range
    With COMMODITIES
        Set commodityRange = .Cells(.Cells(.Rows.Count, dateCol).End(xlUp).Row, 1)
        Set commodityRange = commodityRange.offset(-(commodityRange.Row - commodityRange.End(xlUp).Row), 0).Resize(commodityRange.Row - commodityRange.End(xlUp).Row + 1, 1)
    End With
    
    Dim CommodityCount As Integer
    CommodityCount = commodityRange.Cells.Count
    Dim priceLastRowNum As Long
    priceLastRowNum = TEFAS_PRICES.Cells(TEFAS_PRICES.Rows.Count, dateCol).End(xlUp).Row
    
    Dim dateDifference As Integer
    dateDifference = CInt(Date - TEFAS_PRICES.Cells(priceLastRowNum, dateCol))
    
    If dateDifference = 0 Then Exit Sub
    
    TEFAS_PRICES.Cells(priceLastRowNum, dateCol).offset(1, 0).Resize(CommodityCount, 1).value = TEFAS_PRICES.Cells(TEFAS_PRICES.Rows.Count, dateCol).End(xlUp).value + 1
    TEFAS_PRICES.Cells(priceLastRowNum, currencyCol).offset(1, 0).Resize(CommodityCount, 1).Value2 = commodityRange.Cells.Value2
    Dim i As Integer
        For i = 0 To (CommodityCount - 1)
            TEFAS_PRICES.Cells(priceLastRowNum, rateCol).offset(1 + i, 0).Value2 = PriceThat(TEFAS_PRICES.Cells(priceLastRowNum, rateCol).offset(1 + i, 0), _
                                                                                            TEFAS_PRICES.Cells(priceLastRowNum, currencyCol).offset(1 + i, 0), _
                                                                                            TEFAS_PRICES.Cells(priceLastRowNum, dateCol).offset(1 + i, 0))
    Next i

    If dateDifference <> 0 Then Call GetNewPriceData
    
End Sub
Private Sub ExportHledgerFile()

    Dim rownum As Long
    Dim splitRowNum As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim myFileAddr As String
    Dim tempFileAddr As String
    Dim portfolioCsvPath As String
    Dim portfolioCsvPath_Sws As String 'portfolioData path for simplywallst web page
    Dim portfolioCsvPath_investing As String
    Dim portfolioCashCsvPath As String
    tempFileAddr = "C:\Budgeting\DATA\Temp.txt"
    myFileAddr = "C:\Budgeting\DATA\Main.txt"
    portfolioCsvPath = "C:\Budgeting\DATA\PortfolioMovements.csv"
    portfolioCsvPath_Sws = "C:\Budgeting\DATA\PortfolioMovements_SimplyWallSt.csv"
    portfolioCashCsvPath = "C:\Budgeting\DATA\PortfolioCashMovements.csv"
    portfolioCsvPath_investing = "C:\Budgeting\DATA\PortfolioMovements_OpenPositions.csv"
    Dim hledgerFile As Object
    Set hledgerFile = fso.OpenTextFile(tempFileAddr, 2, True, -2)

    hledgerFile.WriteLine "; created by me at " & Now()
    hledgerFile.WriteLine
    hledgerFile.WriteLine "; Options"
    hledgerFile.WriteLine "decimal-mark ,"
    hledgerFile.WriteLine "commodity 1.000,00 TRY"
    hledgerFile.WriteLine
    hledgerFile.WriteLine "; Special Accounts"
    hledgerFile.WriteLine "account Varlýklar                    ; type: A"
    hledgerFile.WriteLine "account Borçlar                      ; type: L"
    hledgerFile.WriteLine "account Özkaynaklar                  ; type: E"
    hledgerFile.WriteLine "account Gelir                        ; type: R"
    'hledgerFile.WriteLine "account Gelir:Yatýrým                ; type: R" 'for special purposes
    hledgerFile.WriteLine "account Gider                        ; type: X"
    hledgerFile.WriteLine "account Varlýklar:Dönen Varlýklar    ; type: C"
    hledgerFile.WriteLine
    
    Dim Accountes As Object
    Dim anAccount As Variant
    Dim longestAccountNameLen As Long
    Set Accountes = GetTransactionAccountNamess
    ACCOUNTS.Cells.Clear
    For Each anAccount In Accountes
        ACCOUNTS.Cells(ACCOUNTS.Rows.Count, 1).End(xlUp).offset(1, 0).value = anAccount
    Next anAccount
    With ACCOUNTS.Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    hledgerFile.WriteLine "; Accounts"
    For Each anAccount In Accountes
        hledgerFile.WriteLine "account " & anAccount
        If Len(anAccount) > longestAccountNameLen Then longestAccountNameLen = Len(anAccount)
    Next anAccount
    hledgerFile.WriteLine
    
    Dim entities As Object
    Dim entity As Variant
    Set entities = GetTransactionMinMaxEntityDates
    COMMODITIES.Cells.Clear
    For Each entity In entities
        If entities(entity)(0) <> 0 Then COMMODITIES.Cells(COMMODITIES.Rows.Count, 1).End(xlUp).offset(1, 0).value = entity
    Next entity
    With COMMODITIES.Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    hledgerFile.WriteLine "; Commodities"
    For Each entity In entities
        'Currency isminde karakterin sayýsal - number olmasýna karþý önlem "TI3" vakasý.
        With CreateObject("VBScript.RegExp")
            .pattern = "(\d)"
            If .test(entity) Then entity = """" & entity & """"
        End With
        '
        hledgerFile.WriteLine "commodity " & entity & " 1.000,00"
    Next entity
    hledgerFile.WriteLine
    
    'get new commodity prices
    'If MsgBox("Get new prices by 'ComparePrices' module?", vbYesNo, "Prices") = vbYes Then Call ComparePriceList
    If MsgBox("Get new prices by 'GetNewPrices' module?", vbYesNo, "Prices") = vbYes Then Call GetNewPriceData

    Dim ddate As Date
    Dim ddateStr As String
    Dim transCode As String
    Dim Description As String
    Dim account As String
    Dim amount As String
    Dim mainCurrncy As String
    Dim splitCurrncy As String
    Dim commentTransaction As String
    Dim commentSplit1 As String
    Dim commentSplitN As String
    Dim blanks As String
    Dim assertion As String
    Dim profit As String
    
    'Portfolio csv data
    Dim portfolioData As String
    Dim portfolioData_Sws As String 'portfolioData for simplywallst web page
    Dim portfolioCashData As String
    
    Dim commodityCommandsDict As Scripting.Dictionary
    Set commodityCommandsDict = New Scripting.Dictionary
    
    ' > Address column headers
    Call PopulateColumnHeaderIndexes("MAIN_LEDGER")
    ' < Address column headers
    
    rownum = 2
    With MAIN_LEDGER
        Do While .Cells(rownum, dateCol) <> ""
            ' Get transaction values
            ddate = Replace(.Cells(rownum, dateCol).value, ".", "-")
            ddateStr = Year(ddate) & "-" & IIf(Month(ddate) < 10, "0", "") & Month(ddate) & "-" & IIf(Day(ddate) < 10, "0", "") & Day(ddate)
            transCode = "(" & .Cells(rownum, transCodeCol).value & ")"
            Description = .Cells(rownum, descriptionCol).value
            account = .Cells(rownum, accNameCol).value
            amount = .Cells(rownum, amountCol).value
            mainCurrncy = GetCurrency(.Cells(rownum, currencyCol))
            commentTransaction = .Cells(rownum, notesCol).value
            commentSplit1 = .Cells(rownum, tagCol).value
            
            '> Parse transaction data
            'first line
            hledgerFile.WriteLine ddateStr & "  *  " & transCode & "  " & Description '& "  " & IIf(commentTransaction <> "", ";" & commentTransaction, "")
            'first Posting
            blanks = String(longestAccountNameLen - Len(account) + IIf(left(amount, 1) = "-", 0, 1), " ")
            assertion = IIf(.Cells(rownum, reconCol).value = "", "", "=" & .Cells(rownum, reconCol).value & " " & mainCurrncy & "  ")
            hledgerFile.WriteLine "  " & account & blanks & "  " & amount & " " & mainCurrncy & "  " & assertion & IIf(commentSplit1 <> "", ";" & commentSplit1, "")
            
            splitRowNum = rownum + 1
            Do While .Cells(splitRowNum, dateCol) = "" And .Cells(splitRowNum, accNameCol) <> ""
                account = .Cells(splitRowNum, accNameCol).value
                amount = .Cells(splitRowNum, amountCol).value
                splitCurrncy = GetCurrency(.Cells(splitRowNum, currencyCol), mainCurrncy)
                commentSplitN = .Cells(splitRowNum, tagCol).value
                'Nth posting
                blanks = String(longestAccountNameLen - Len(account) + IIf(left(amount, 1) = "-", 0, 1), " ")
                assertion = IIf(.Cells(splitRowNum, reconCol).value = "", "", "=" & .Cells(splitRowNum, reconCol).value & " " & splitCurrncy & "  ")
                hledgerFile.WriteLine "  " & account & blanks & "  " & amount & " " & splitCurrncy & "  " & assertion & IIf(commentSplitN <> "", ";" & commentSplitN, "")
                'lots special part
                If .Cells(splitRowNum, operationCol).value = "Buy" Then
                    portfolioData = portfolioData & ddateStr & ";" & "Buy" & ";" & amount & ";" & Split(splitCurrncy, " @ ")(0) & ";" & Replace(Split(splitCurrncy, " @ ")(1), " TRY", "") & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    If Len(Split(Replace(splitCurrncy, """", ""), " @ ")(0)) <> 3 Then
                        portfolioData_Sws = portfolioData_Sws & ddateStr & ";" & "Buy" & ";" & amount & ";" & Split(splitCurrncy, " @ ")(0) & ";" & Replace(Split(splitCurrncy, " @ ")(1), " TRY", "") & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    Else
                    End If
                    portfolioCashData = portfolioCashData & ddateStr & ";" & "Deposit" & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    
'                    hledgerFile.WriteLine "  " & "[" & "Lot:" & left(splitCurrncy, InStr(splitCurrncy, " ") - 1) & ":D" & Replace(ddateStr, "-", "") & "]  -" & amount & " " & splitCurrncy
'                    hledgerFile.WriteLine "  " & "[" & "Lot:" & left(splitCurrncy, InStr(splitCurrncy, " ") - 1) & ":D" & Replace(ddateStr, "-", "") & "]   " & _
                        CDec(amount) * CDec(left(Mid(splitCurrncy, InStr(splitCurrncy, "@") + 2), InStr(Mid(splitCurrncy, InStr(splitCurrncy, "@") + 2), " ") - 1)) & " " & mainCurrncy
                        
                    commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & amount & "::" & splitCurrncy & "::BUY", 1  'stock command
                    
                ElseIf .Cells(splitRowNum, operationCol).value = "Sell" Then
                    portfolioData = portfolioData & ddateStr & ";" & "Sell" & ";" & amount & ";" & Split(splitCurrncy, " @ ")(0) & ";" & Replace(Split(splitCurrncy, " @ ")(1), " TRY", "") & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    If Len(Split(Replace(splitCurrncy, """", ""), " @ ")(0)) <> 3 Then
                        portfolioData_Sws = portfolioData_Sws & ddateStr & ";" & "Sell" & ";" & amount & ";" & Split(splitCurrncy, " @ ")(0) & ";" & Replace(Split(splitCurrncy, " @ ")(1), " TRY", "") & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    Else
                    End If
                    portfolioCashData = portfolioCashData & ddateStr & ";" & "Removal" & ";" & (CDec(Replace(Split(splitCurrncy, " @ ")(1), " TRY", "")) * (CDec(amount))) & vbCrLf
                    
                    'hledgerFile.WriteLine "  " & "[" & "Lot:" & ddateStr & ":" & amount & " " & splitCurrncy & "]   " & -1 * Cdec(amount) & " " & splitCurrncy
                    'hledgerFile.WriteLine "  " & "[" & "Lot:" & ddateStr & ":" & amount & " " & splitCurrncy & "]   " & _
                        Cdec(amount) * Cdec(Left(Mid(splitCurrncy, InStr(splitCurrncy, "@") + 2), InStr(Mid(splitCurrncy, InStr(splitCurrncy, "@") + 2), " ") - 1)) & " " & mainCurrncy
                    
                    commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & amount & "::" & splitCurrncy & "::SELL", 1 'unstock command
                
                ElseIf left(.Cells(splitRowNum, operationCol).value, 5) = "Split" Then
                    ddate = CDate(Mid(.Cells(rownum, descriptionCol).value, InStr(1, .Cells(rownum, descriptionCol).value, "Tarih:") + 6))
                    ddateStr = Year(ddate) & "-" & IIf(Month(ddate) < 10, "0", "") & Month(ddate) & "-" & IIf(Day(ddate) < 10, "0", "") & Day(ddate)
                    commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & Mid(.Cells(splitRowNum, operationCol).value, 7) & "::" & splitCurrncy & "::SPLIT", 1 'unstock command
                Else
                End If
                splitRowNum = splitRowNum + 1
            Loop
            '< Parse transaction data
            rownum = splitRowNum
            hledgerFile.WriteLine
        Loop
    End With
    
    'Parse Lots
    Dim writeCommands As Scripting.Dictionary
    Set writeCommands = ParseCommodities(commodityCommandsDict)
    'write lots
    Dim aCommand As Variant
    Dim jumper As Long
    Dim i As Long
    Dim streamLine As String
    hledgerFile.Close
    Dim tempReadFile As Object
    Set tempReadFile = fso.OpenTextFile(tempFileAddr, 1, False, -2)
    Set hledgerFile = fso.OpenTextFile(myFileAddr, 2, True, -2) 'to go to the begining of the file
    
    'get to the position and change writing according to buy sell stuff.
    Dim leftCountTemp As Integer
    Dim commandPartArray() As String
    Dim commandPartNum As Integer
    Dim tempString As String
    Do While Not tempReadFile.AtEndOfStream
        streamLine = tempReadFile.ReadLine
        '
        If writeCommands.Exists(CStr(tempReadFile.line)) Then
            If Split(writeCommands(CStr(tempReadFile.line)), "|")(1) Like "BUY*" Then
                writeCommands(CStr(tempReadFile.line)) = Split(writeCommands(CStr(tempReadFile.line)), "|")(0) & vbCrLf
                leftCountTemp = Len(streamLine)
                Do While Mid(streamLine, leftCountTemp - 2, 3) <> "   "
                    leftCountTemp = leftCountTemp - 1
                Loop
                hledgerFile.WriteLine left(streamLine, leftCountTemp) & writeCommands(CStr(tempReadFile.line))
            Else 'sell command
                For i = LBound(Split(writeCommands(CStr(tempReadFile.line)), "|SELL")) To UBound(Split(writeCommands(CStr(tempReadFile.line)), "|SELL")) - 1
                tempString = Split(writeCommands(CStr(tempReadFile.line)), "|SELL" & vbCrLf)(i) & vbCrLf
                'hledgerFile.WriteLine streamLine
                'streamLine = tempReadFile.ReadLine
                leftCountTemp = Len(streamLine)
                Do While Mid(streamLine, leftCountTemp - 2, 3) <> "   "
                    leftCountTemp = leftCountTemp - 1
                Loop
                commandPartArray = Split(tempString, vbCrLf)
                For commandPartNum = 0 To ((LineCounter(tempString) \ 4) - 1)
                    hledgerFile.WriteLine left(streamLine, leftCountTemp) & commandPartArray((commandPartNum - 0) * 4)
                    hledgerFile.WriteLine commandPartArray(((commandPartNum - 0) * 4) + 1)
                    hledgerFile.WriteLine commandPartArray(((commandPartNum - 0) * 4) + 2)
                    hledgerFile.WriteLine commandPartArray(((commandPartNum - 0) * 4) + 3)
                Next commandPartNum
                jumper = jumper + 3
                Next i
                hledgerFile.WriteLine "  Gelir:Yatýrým"
            End If
        Else
            hledgerFile.WriteLine streamLine
        End If
    Loop
    
    'permission denied
    'fso.DeleteFile tempFileAddr, True
    
    ' > Address column headers
    Call PopulateColumnHeaderIndexes("TEFAS_PRICES")
    ' < Address column headers
    
    rownum = 2
    hledgerFile.WriteLine
    With TEFAS_PRICES
        Do While .Cells(rownum, dateCol) <> ""
            ddate = Replace(.Cells(rownum, dateCol).value, ".", "-")
            ddateStr = Year(ddate) & "-" & IIf(Month(ddate) < 10, "0", "") & Month(ddate) & "-" & IIf(Day(ddate) < 10, "0", "") & Day(ddate)
            amount = .Cells(rownum, rateCol).value
            mainCurrncy = .Cells(rownum, currencyCol)
            'Currency isminde karakterin sayýsal - number olmasýna karþý önlem "TI3" vakasý.
            With CreateObject("VBScript.RegExp")
                .pattern = "(\d)"
                If .test(mainCurrncy) Then mainCurrncy = """" & mainCurrncy & """"
            End With
            '
            hledgerFile.WriteLine "P    " & ddateStr & "    " & mainCurrncy & "    " & amount & " TRY"
            rownum = rownum + 1
        Loop
    End With
 
    'Create portfolioCsvPath for portfolio-performance
    Dim line As String
    Dim cashMovementsLine As String
    Dim tempLine As Variant
    Dim fileNo As Variant
    Dim stockObj As Variant
    Dim commodityCommandsDictDates As Variant
    Set commodityCommandsDictDates = New Scripting.Dictionary
    For i = 0 To commodityCommandsDict.Count - 1
        line = commodityCommandsDict.keys(i)
        commodityCommandsDictDates.item(Split(line, "::")(0)) = Split(line, "::")(1)
    Next i
    line = ""
    cashMovementsLine = ""
    fileNo = FreeFile
    Open portfolioCsvPath For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    For i = 0 To writeCommands.Count - 1
        For Each tempLine In Split(writeCommands.Items(i), vbCrLf)
            Do While left(tempLine, 1) = " "
                tempLine = Mid(tempLine, 2)
            Loop
            If left(tempLine, 1) <> "[" And tempLine <> "" Then
                line = line & commodityCommandsDictDates(writeCommands.keys(i))
                stockObj = Split(tempLine, " ")
                line = line & ";" & IIf(stockObj(0) > 0, "Buy", "Sell") & ";" & stockObj(0)
                line = line & ";" & stockObj(1) & ";" & Abs(CDbl(stockObj(3))) & ";" & stockObj(0) * Abs(CDbl(stockObj(3)))
                line = line & vbCrLf
                cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                cashMovementsLine = cashMovementsLine & ";" & IIf(CDbl(stockObj(0)) < 0, "Removal", "Deposit") & ";" & stockObj(0) * stockObj(3)
                cashMovementsLine = cashMovementsLine & vbCrLf
            Else
            End If
        Next tempLine
    Next i
    line = Replace(line, """", "")
    cashMovementsLine = Replace(cashMovementsLine, """", "")
    Print #fileNo, line 'portfolioData
    Close #fileNo
    
'    fileNo = FreeFile
'    Open portfolioCsvPath_Sws For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
'    Print #fileNo, portfolioData_Sws
'    Close #fileNo
    
    fileNo = FreeFile
    Open portfolioCashCsvPath For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    Print #fileNo, cashMovementsLine
    Close #fileNo

    fileNo = FreeFile
    Open portfolioCsvPath_investing For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    Dim stockObjSub As Variant
    Dim stockObjSubSub As Variant
    For Each stockObj In commDict.keys
        For Each stockObjSub In commDict(stockObj).keys
            For Each stockObjSubSub In commDict(stockObj)(stockObjSub).keys
                line = line & stockObj & ";" & Mid(stockObjSub, 7, 2) & "." & Mid(stockObjSub, 5, 2) & "." & Mid(stockObjSub, 1, 4) & _
                    ";" & Split(commDict(stockObj)(stockObjSub)(stockObjSubSub), "@")(0) & _
                    ";" & Format(Split(commDict(stockObj)(stockObjSub)(stockObjSubSub), "@")(1), "0.0#######") & vbCrLf
            Next stockObjSubSub
        Next stockObjSub
    Next stockObj
    line = Replace(line, """", "")
    Print #fileNo, line
    Close #fileNo
    
    hledgerFile.Close
    Call convertTxttoUTF(myFileAddr, myFileAddr)
    Set tempReadFile = Nothing
    fso.DeleteFile tempFileAddr, True

End Sub
Private Function LineCounter(str As String) As Integer
    LineCounter = (Len(str) - Len(Replace(str, vbCrLf, ""))) * 0.5
End Function
Private Function ParseCommodities(ByRef commodityCommandsDict As Object) As Object

Dim command As Variant
Dim i As Long
Dim commandLineNum() As String
ReDim commandLineNum(1 To commodityCommandsDict.Count)

Dim commands() As Variant
commands = commodityCommandsDict.keys()

Dim commName As String
Dim commPrice As Variant
Dim commDateStamp As String
Dim commQuantity As Variant
Dim tempString() As String
Dim isBuy As Boolean
Dim commLineNum As String
Dim minDateStamp As String
Set commDict = New Scripting.Dictionary
Dim aStock As Scripting.Dictionary
Dim aDate As Scripting.Dictionary
Dim buyedPrice As Variant
Dim buyedQuant As Variant
Dim minSameDayNumber As Long
Dim keyy As Variant

Dim writeCommands As Scripting.Dictionary
Set writeCommands = New Scripting.Dictionary
Dim splitCommands As Scripting.Dictionary
Set splitCommands = New Scripting.Dictionary

'get split commands
For i = UBound(commands) To LBound(commands) Step -1
    tempString = Split(commands(i), "::")
    If tempString(4) = "SPLIT" Then
        commName = left(tempString(3), InStr(tempString(3), " ") - 1)
        commDateStamp = Replace(tempString(1), "-", "")
        commQuantity = CDec(tempString(2))
        splitCommands.item(commDateStamp & "|" & commName) = commQuantity
    Else
    End If
Next i

'parse commodity splits
Dim aSplit As Variant
Dim splitPerc As Double
Dim tempStringPart As Variant
Dim tempVal As String
For Each aSplit In splitCommands
    commName = Split(aSplit, "|")(1)
    commDateStamp = Split(aSplit, "|")(0)
    splitPerc = (CDbl(splitCommands(aSplit)) / 100) + 1
    For i = UBound(commands) To LBound(commands) Step -1
        tempString = Split(commands(i), "::")
        If left(tempString(3), InStr(1, tempString(3), " ") - 1) = commName _
            And CLng(Replace(tempString(1), "-", "")) <= CLng(commDateStamp) _
            And tempString(4) <> "SPLIT" Then
            tempString(0) = CStr(CLng(tempString(0)) - 0)
            tempString(2) = CStr(CDbl(tempString(2)) * splitPerc)
            tempStringPart = tempString(3)
            tempStringPart = Mid(tempStringPart, InStr(1, tempStringPart, "@") + 2)
            tempStringPart = left(tempStringPart, InStr(1, tempStringPart, " ") - 1)
            tempString(3) = Replace(tempString(3), tempStringPart, CStr(CDbl(tempStringPart) / splitPerc))
            commands(i) = ""
            For Each tempStringPart In tempString
                commands(i) = commands(i) & "::" & tempStringPart
            Next tempStringPart
            commands(i) = Mid(commands(i), 3)
        Else
        End If
    Next i
Next aSplit

For i = UBound(commands) To LBound(commands) Step -1
    tempString = Split(commands(i), "::")
    commLineNum = tempString(0)
    'If commLineNum = "15683" Then Stop
    isBuy = IIf(tempString(4) = "BUY", True, False)
    commName = left(tempString(3), InStr(tempString(3), " ") - 1)
    commQuantity = CDec(tempString(2))
    commDateStamp = Replace(tempString(1), "-", "")
    tempString(3) = Mid(tempString(3), InStr(tempString(3), "@") + 2)
    tempString(3) = left(tempString(3), InStr(tempString(3), " ") - 1)
    commPrice = Format(CDec(tempString(3)), "0.0#########################################")
    'If commName = "TKM" Then Stop
    'If commLineNum = "2673" Then Stop
    If commDict.Exists(commName) Then
        If IIf(tempString(4) = "BUY", True, False) Then
            Set aStock = commDict(commName)
            If aStock.Exists(commDateStamp) Then
                Set aDate = aStock(commDateStamp)
                minSameDayNumber = 0
                For Each keyy In aDate.keys
                    If keyy > minSameDayNumber Then minSameDayNumber = keyy 'actually it is the largest sameday num @TODO
                Next keyy
                aDate.Add minSameDayNumber + 1, commQuantity & "@" & commPrice
            Else
                Set aDate = New Scripting.Dictionary
                aDate.Add 1, commQuantity & "@" & commPrice
                Set aStock.item(commDateStamp) = aDate
            End If
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                "   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " TRY" & "|BUY" & vbCrLf
        ElseIf IIf(tempString(4) = "SELL", True, False) Then
            Set aStock = commDict(commName)
            Do
                minDateStamp = "99999999"
                For Each keyy In aStock.keys
                    If keyy < minDateStamp Then minDateStamp = keyy
                Next keyy
                minSameDayNumber = 99999
                For Each keyy In aStock(minDateStamp).keys
                    If keyy < minSameDayNumber Then minSameDayNumber = keyy
                Next keyy
                tempString = Split(aStock(minDateStamp).item(minSameDayNumber), "@")
                buyedPrice = CDec(tempString(1))
                buyedQuant = CDec(tempString(0))
                'SELLING OPTIONS
                'If commName = "TKM" Then Stop
                If Abs(commQuantity) > buyedQuant Then
                    aStock(minDateStamp).remove (minSameDayNumber)
                    If aStock(minDateStamp).Count = 0 Then aStock.remove (minDateStamp)
                    If aStock.Count = 0 Then commDict.remove (commName)
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(buyedQuant) & " " & commName & " @ " & Format(buyedPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & Abs(buyedQuant) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & -1 * buyedQuant * buyedPrice & " TRY" & vbCrLf & _
                                                      "  [Lot:Profit:" & commName & "]   " & -1 * (Abs(buyedQuant) * (commPrice - buyedPrice)) & " TRY" & "|SELL" & vbCrLf
                    commQuantity = commQuantity + buyedQuant
                ElseIf Abs(commQuantity) < buyedQuant Then
                    aStock(minDateStamp).item(minSameDayNumber) = buyedQuant - Abs(commQuantity) & "@" & buyedPrice
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(commQuantity) & " " & commName & " @ " & Format(buyedPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & commQuantity * buyedPrice & " TRY" & vbCrLf & _
                                                      "  [Lot:Profit:" & commName & "]   " & -1 * (Abs(commQuantity) * (commPrice - buyedPrice)) & " TRY" & "|SELL" & vbCrLf
                    commQuantity = 0
                Else
                    aStock(minDateStamp).remove (minSameDayNumber)
                    If aStock(minDateStamp).Count = 0 Then aStock.remove (minDateStamp)
                    If aStock.Count = 0 Then commDict.remove (commName)
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(commQuantity) & " " & commName & " @ " & Format(buyedPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " TRY" & vbCrLf & _
                                                      "  [Lot:" & commName & ":D" & commDateStamp & "]   " & commQuantity * buyedPrice & " TRY" & vbCrLf & _
                                                      "  [Lot:Profit:" & commName & "]   " & -1 * (Abs(commQuantity) * (commPrice - buyedPrice)) & " TRY" & "|SELL" & vbCrLf
                    commQuantity = 0
                End If
                'END OF SELLING OPTIONS
            Loop While commQuantity <> 0
        ElseIf IIf(tempString(4) = "SPLIT", True, False) Then
            '
        Else
        End If
        
    Else
        If isBuy Then
            Set aStock = New Scripting.Dictionary
            Set aDate = New Scripting.Dictionary
            aDate.Add 1, commQuantity & "@" & commPrice
            aStock.Add commDateStamp, aDate
            commDict.Add commName, aStock
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                "   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " TRY" & "|BUY" & vbCrLf
        Else
            MsgBox ("err-1")
        End If
    End If
Next i

Set ParseCommodities = writeCommands

End Function

Private Function GetCurrency(ByVal rng As Range, Optional ByVal mainCurrency As String) As String

    If rng.value Like "CURRENCY::*" Then
        GetCurrency = Replace(rng.value, "CURRENCY::", "")
    ElseIf rng.offset(0, 1).value = "Buy" Or rng.offset(0, 1).value = "Sell" Or left(rng.offset(0, 1).value, 5) = "Split" Then
        GetCurrency = rng.offset(0, 3).value
        Do Until InStr(GetCurrency, ":") = 0
            GetCurrency = Mid(GetCurrency, InStr(GetCurrency, ":") + 1)
        Loop
        'Currency isminde karakterin sayýsal - number olmasýna karþý önlem "TI3" vakasý.
        With CreateObject("VBScript.RegExp")
            .pattern = "(\d)"
            'If GetCurrency = "A1CAP" Then Stop
            If .test(GetCurrency) Then GetCurrency = """" & GetCurrency & """"
        End With
        '
        GetCurrency = GetCurrency & " @ " & Format(rng.offset(0, 5).value, "0.0#######################################") & " " & mainCurrency 'No scientific for you!!
    Else
        Do While GetCurrency = ""
            GetCurrency = GetCurrency(rng.offset(-1, 0))
            Set rng = rng.offset(-1, 0)
        Loop
    End If

End Function

Private Sub ExportAsCSV()

    Dim MyFileName As String
    Dim CurrentWB As Workbook, TempWB As Workbook

    Set CurrentWB = ActiveWorkbook
    ActiveWorkbook.ActiveSheet.UsedRange.Copy

    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With

    'MyFileName = CurrentWB.Path & "\" & Left(CurrentWB.Name, InStrRev(CurrentWB.Name, ".") - 1) & ".csv"
    'Optionally, comment previous line and uncomment next one to save as the current sheet name
    MyFileName = CurrentWB.path & "\" & CurrentWB.ActiveSheet.name & ".csv"


    Application.DisplayAlerts = False
    TempWB.SaveAs fileName:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Call convertTxttoUTF(MyFileName, MyFileName)
    
End Sub

Private Sub convertTxttoUTF(sInFilePath As String, sOutFilePath As String)
    Dim objFS  As Object
    Dim objFSNoBOM  As Object
    Dim iFile       As Variant
    Dim sFileData   As String
    
    'Init
    iFile = FreeFile
    Open sInFilePath For Input As #iFile
        sFileData = Input$(LOF(iFile), iFile)
        sFileData = sFileData & vbCrLf
    Close iFile

    'Open & Write
    Set objFS = CreateObject("ADODB.Stream")
    With objFS
        .Charset = "UTF-8"
        .Open
        .WriteText sFileData
        .Position = 0
        .Type = 2
        .Position = 3
    End With
    Set objFSNoBOM = CreateObject("ADODB.Stream")
    With objFSNoBOM
        .Type = 1
        .Open
        objFS.CopyTo objFSNoBOM
    End With
        
    'Save & Close
    objFSNoBOM.SaveToFile sOutFilePath, 2   '2: Create Or Update
    objFSNoBOM.Close
    objFS.Close
    
    'Completed
    Application.StatusBar = "Completed"
End Sub

Public Sub ComparePriceList()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim transactionDates As Object
    Set transactionDates = GetTransactionMinMaxEntityDates
    Dim priceDates As Object
    Set priceDates = GetPricesMinMaxEntityDates
    Dim anEntityName As Variant
    Dim pricesToWrite As Variant
    pricesToWrite = 0
    
    For Each anEntityName In transactionDates.keys
        'check if exits
        'If anEntityName = "ZPLIB" Then Stop
        If priceDates.Exists(anEntityName) = False Then
            pricesToWrite = PriceThat(CStr(anEntityName), , CDate(transactionDates(anEntityName)(1)), CDate(transactionDates(anEntityName)(2)))
            Call WritePrices(pricesToWrite, anEntityName)
            pricesToWrite = 0
        ElseIf TypeName(priceDates(anEntityName)) <> "Variant()" Then 'EmptyArray
            pricesToWrite = PriceThat(CStr(anEntityName), , CDate(transactionDates(anEntityName)(1)), CDate(transactionDates(anEntityName)(2)))
            Call WritePrices(pricesToWrite, anEntityName)
            pricesToWrite = 0
        Else
            'compare min dates
            If transactionDates(anEntityName)(1) < priceDates(anEntityName)(1) Then
                pricesToWrite = PriceThat(CStr(anEntityName), , CDate(transactionDates(anEntityName)(1)), CDate(priceDates(anEntityName)(1)) - 1)
            ElseIf transactionDates(anEntityName)(1) > priceDates(anEntityName)(1) Then
                
            Else
            End If
            Call WritePrices(pricesToWrite, anEntityName)
            pricesToWrite = 0
            'compare max dates
            If transactionDates(anEntityName)(2) < priceDates(anEntityName)(2) Then
                
            ElseIf transactionDates(anEntityName)(2) > priceDates(anEntityName)(2) Then
                pricesToWrite = PriceThat(CStr(anEntityName), , CDate(priceDates(anEntityName)(2)) + 1, CDate(transactionDates(anEntityName)(2)))
            Else
            End If
            Call WritePrices(pricesToWrite, anEntityName)
            pricesToWrite = 0
            'compare current day for stocks (get current price for today)
'            If CDate(transactionDates(anEntityName)(2)) = Date And _
'               CDate(priceDates(anEntityName)(2)) <> Date And _
'               (Len(anEntityName) > 3 Or anEntityName = "USD" Or anEntityName = "EUR") Then
'                ReDim pricesToWrite(0, 0 To 1)
'                pricesToWrite(0, 0) = CDate(transactionDates(anEntityName)(2))
'                pricesToWrite(0, 1) = PriceThat(CStr(anEntityName), , CDate(transactionDates(anEntityName)(2)))
'            Else
'            End If
'            Call WritePrices(pricesToWrite, anEntityName)
'            pricesToWrite = 0
        End If
    Next anEntityName
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Public Sub WritePrices(pricesToWrite As Variant, entityName As Variant)
    
    Dim i As Long
    Dim startRow As Long
    With TEFAS_PRICES
        startRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        If TypeName(pricesToWrite) = "Variant()" Then
            For i = 0 To UBound(pricesToWrite)
                .Cells(1, 1).offset(startRow + i - 1, 0).value = pricesToWrite(i, 0)
                .Cells(1, 1).offset(startRow + i - 1, 1).value = entityName
                .Cells(1, 1).offset(startRow + i - 1, 2).value = pricesToWrite(i, 1)
            Next i
        Else
            If pricesToWrite <> 0 Then
                .Cells(1, 1).offset(startRow + 0 - 1, 0).value = Date
                .Cells(1, 1).offset(startRow + 0 - 1, 1).value = entityName
                .Cells(1, 1).offset(startRow + 0 - 1, 2).value = pricesToWrite
            Else
            
            End If
        End If
    End With
End Sub

Private Function GetTransactionMinMaxEntityDates() As Object
      
    Dim rownum As Long
    Dim splitRowNum As Long
    rownum = 2 'starting row
    Dim entityDict As Object
    Set entityDict = CreateObject("scripting.Dictionary")
    Dim entityName As String
    Dim ddate As Long 'transaction date
    Dim dates(2) As Variant 'dates and count data
    Dim entityCount As Long 'entity transaction count
    With MAIN_LEDGER
        Do While .Cells(rownum, 1) <> ""
            ddate = .Cells(rownum, 1)
            splitRowNum = rownum + 1
            Do While .Cells(splitRowNum, 1) = "" And .Cells(splitRowNum, 8) <> ""
                If .Cells(splitRowNum, 6) = "Buy" Or .Cells(splitRowNum, 6) = "Sell" Then
                    entityCount = .Cells(splitRowNum, 9)
                    entityName = .Cells(splitRowNum, 8)
                    Do While InStr(entityName, ":") <> 0
                        entityName = Mid(entityName, InStr(entityName, ":") + 1)
                    Loop
                    'If left(entityName, 2) = "W_" Then Stop
                    'If ddate = 44796 Then Stop
                    If entityDict.Exists(entityName) Then
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityDict(entityName)(0) + entityCount
                        dates(1) = entityDict(entityName)(1)
                        dates(2) = entityDict(entityName)(2)
                    Else
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityCount
                        dates(1) = CDate(Date)
                        dates(2) = CDate(25569) '1.1.1970 in long format
                    End If
                    If ddate < dates(1) Then dates(1) = CDate(ddate)
                    If ddate > dates(2) Then dates(2) = CDate(ddate)
                    entityDict.item(entityName) = dates
                Else
                
                End If
                splitRowNum = splitRowNum + 1
            Loop
            rownum = splitRowNum
        Loop
        Debug.Print rownum
    End With
    Dim entityKey As Variant
    Dim i As Integer
    Dim entArray As Variant
    For Each entityKey In entityDict.keys
        'If entityKey = "GARFA" Then Stop
        i = i + 1
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 0).value = entityKey
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 1).value = entityDict(entityKey)(0)
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 2).value = entityDict(entityKey)(1)
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 3).value = CDate(IIf(entityDict(entityKey)(0) = 0, entityDict(entityKey)(2), CLng(Date)))
        If entityDict(entityKey)(0) <> 0 Then
            entArray = entityDict(entityKey)
            entArray(2) = Date
            entityDict(entityKey) = entArray
        Else
        End If
    Next entityKey
    
    Set GetTransactionMinMaxEntityDates = entityDict
    
End Function
    

Private Function GetPricesMinMaxEntityDates()

    Dim rownum As Long
    rownum = 2 'starting row
    Dim entityDict As Object
    Set entityDict = CreateObject("scripting.Dictionary")
    Dim entityName As String
    Dim ddate As Long 'transaction date
    Dim dates(2) As Variant 'dates and count data
    Dim entityCount As Long 'entity transaction count
    With TEFAS_PRICES
        Do While .Cells(rownum, 1) <> ""
            ddate = .Cells(rownum, 1)
            entityCount = 0
            entityName = .Cells(rownum, 2)
                    'If ddate = 44796 Then Stop
                    If entityDict.Exists(entityName) Then
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityDict(entityName)(0) + entityCount
                        dates(1) = entityDict(entityName)(1)
                        dates(2) = entityDict(entityName)(2)
                    Else
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityCount
                        dates(1) = Date
                        dates(2) = CDate(25569) '1.1.1970 in long format
                    End If
                    If ddate < dates(1) Then dates(1) = CDate(ddate)
                    If ddate > dates(2) Then dates(2) = CDate(ddate)
                    entityDict.item(entityName) = dates
            rownum = rownum + 1
        Loop
    End With
'    Dim entityKey As Variant
'    Dim i As Long
'    For Each entityKey In entityDict.keys
'        i = i + 1
'        ActiveSheet.Cells(1, 6).End(xlUp).offset(i, 0).value = entityKey
'        ActiveSheet.Cells(1, 6).End(xlUp).offset(i, 1).value = entityDict(entityKey)(0)
'        ActiveSheet.Cells(1, 6).End(xlUp).offset(i, 2).value = entityDict(entityKey)(1)
'        ActiveSheet.Cells(1, 6).End(xlUp).offset(i, 3).value = entityDict(entityKey)(2)
'    Next entityKey

    Set GetPricesMinMaxEntityDates = entityDict

End Function

Private Function GetTransactionAccountNamess() As Object
      
    Dim rownum As Integer
    Dim splitRowNum As Integer
    rownum = 2 'starting row
    Dim accountDict As Object
    Set accountDict = CreateObject("scripting.Dictionary")
    Dim accountName As String
    
    With MAIN_LEDGER
        Do While .Cells(rownum, 1) <> ""
            splitRowNum = rownum + 1
            accountName = .Cells(rownum, 8)
            accountDict.item(accountName) = 1
            Do While .Cells(splitRowNum, 1) = "" And .Cells(splitRowNum, 8) <> ""
                accountName = .Cells(splitRowNum, 8)
                accountDict.item(accountName) = 1
                splitRowNum = splitRowNum + 1
            Loop
            rownum = splitRowNum
        Loop
    End With
    
    Set GetTransactionAccountNamess = accountDict

End Function



