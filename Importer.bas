Attribute VB_Name = "Importer"
Option Explicit

Public Sub ImporterBegin()

    Dim WS As Worksheet
    Set WS = IMPORT_SH
    WS.activate
    
    Dim targetWs As Worksheet
    If WS.Cells(2, 1).value = "" Then Exit Sub
    On Error GoTo WRONGPAGENAME
    Set targetWs = ActiveWorkbook.Worksheets(WS.Cells(2, 1).value)
    On Error GoTo 0
    
    Dim datesRange As Range
    Dim notesRange As Range
    Dim amountRange As Range
    Dim expenseCategoryRange As Range
    Set datesRange = Application.InputBox("Select date columns first value", "Date Column", , , , , , 8)
    Set notesRange = Application.InputBox("Select note columns first value", "Note Column", , , , , , 8)
    Set amountRange = Application.InputBox("Select amount columns first value", "Amount Column", , , , , , 8)
    ' Check selected Ranges
    If datesRange.Cells.Count <> 1 Then GoTo WRONGDATACOUNT
    If notesRange.Cells.Count <> 1 Then GoTo WRONGDATACOUNT
    If amountRange.Cells.Count <> 1 Then GoTo WRONGDATACOUNT
    ' Resize selection to get data
    Set expenseCategoryRange = WS.Cells(amountRange.Row, Application.WorksheetFunction.max(datesRange.Column, notesRange.Column, amountRange.Column) + 1)
    Set datesRange = datesRange.Resize(WS.Cells(WS.Rows.Count, datesRange.Column).End(xlUp).Row - datesRange.Row + 1, 1)
    Set notesRange = notesRange.Resize(WS.Cells(WS.Rows.Count, notesRange.Column).End(xlUp).Row - notesRange.Row + 1, 1)
    Set amountRange = amountRange.Resize(WS.Cells(WS.Rows.Count, amountRange.Column).End(xlUp).Row - amountRange.Row + 1, 1)
    Set expenseCategoryRange = expenseCategoryRange.Resize(amountRange.Rows.Count, 1)
    ' Check data
    If Not (datesRange.Cells.Count = notesRange.Cells.Count And datesRange.Cells.Count = amountRange.Cells.Count) Then GoTo WRONGDATACOUNT
    '*********************
    'check descriptions
    '*********************
    Dim foundDescRange As Range
    Dim answer As Variant
    Dim i As Integer
    Load FrmDescription
    For i = datesRange.Rows.Count To 1 Step -1
        'Check duplicates!
        If CheckDuplicate(datesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs) = 0 And _
                                                                        amountRange.Cells(i, 1).offset(0, 1).value = "" Then
        '
               
                
            Set foundDescRange = Nothing
            'List search
            
            
            
            
            'Pessimistic search
            
            
            
            'Optimistic search
            
            On Error Resume Next
            Set foundDescRange = targetWs.Cells(1, 3).EntireColumn.Find(notesRange.Cells(i, 1).value, LookAt:=xlPart)
            On Error GoTo 0
            Do While Not foundDescRange Is Nothing
            'ask is that it?
                FrmDescription.lblQuestionText.Caption = "is this description appropriate?" & vbCrLf & _
                                                            notesRange.Cells(i, 1).value & vbCrLf & _
                                                            foundDescRange.value & vbCrLf & _
                                                            foundDescRange.offset(1, 5).value
                FrmDescription.show
                If FrmDescription.frmAnswer = "Yes" Then
                    notesRange.Cells(i, 1).value = foundDescRange.value
                    expenseCategoryRange.Cells(i, 1).value = foundDescRange.offset(1, 5).value
                    Set foundDescRange = Nothing
                ElseIf FrmDescription.frmAnswer = "No" Then
                    Set foundDescRange = targetWs.Cells(1, 3).EntireColumn.FindNext(foundDescRange)
                ElseIf FrmDescription.frmAnswer = "Cancel" Then
                    Set foundDescRange = Nothing
                Else 'Edit Mode
                    notesRange.Cells(i, 1).value = FrmDescription.frmAnswer
                    '@TODO expenseCategoryRange.Cells(i, 1).Value = foundDescRange.Offset(1, 5).Value
                    Set foundDescRange = Nothing
                End If
            Loop
        Else
            'duplicate entry found
            'Debug.Print "Duplicate found"
            'datesRange.Cells(i, 1).Interior.ColorIndex = 3
            'amountRange.Cells(i, 1).Interior.ColorIndex = 3
        End If
    Next i
    Unload FrmDescription
        
    'shall we continue to add data
    
    
    
    '********************************
    'Make way for and add new data
    '********************************
    Dim startRow As Long
    Dim reconcileNoteRow As Long
    For i = datesRange.Rows.Count To 1 Step -1
        'Check duplicates!
        If CheckDuplicate(datesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs) = 0 Then
        '            '> find start row for yourself
            startRow = 2 'by default it is 2
            Do Until datesRange.Cells(i, 1).value >= targetWs.Cells(startRow, 1).value And targetWs.Cells(startRow, 1).value <> ""
                startRow = startRow + 1
            Loop
            '< find start row for yourself
            targetWs.Cells(startRow, 1).EntireRow.Insert
            targetWs.Cells(startRow, 1).EntireRow.Insert
            targetWs.Cells(startRow, 1).value = datesRange.Cells(i, 1).value
            targetWs.Cells(startRow, 2).value = "!" 'Random id generator??
            targetWs.Cells(startRow, 3).value = notesRange.Cells(i, 1).value
            targetWs.Cells(startRow, 5).value = "CURRENCY::TRY"
            targetWs.Cells(startRow, 8).value = targetWs.Cells(startRow + 2, 8).value '@TODO decoupling with a dictionary would be fine
            targetWs.Cells(startRow + 1, 8).value = expenseCategoryRange.Cells(i, 1).value
            targetWs.Cells(startRow, 9).value = amountRange.Cells(i, 1).value
            '> check for commodity transaction. If it is then you have to use somethings...
            If datesRange.Cells(i, 1).Interior.ColorIndex <> -4142 Then 'commodity transaction
                targetWs.Cells(startRow + 1, 9).value = IIf(amountRange.Cells(i, 1).value < 0, 1, -1) * _
                                                                        CommodityCount(notesRange.Cells(i, 1).value)
                targetWs.Cells(startRow + 1, 10).value = -1 * (amountRange.Cells(i, 1).value / targetWs.Cells(startRow + 1, 9))
                targetWs.Cells(startRow + 1, 6).value = IIf(amountRange.Cells(i, 1).value < 0, "Buy", "Sell")
            Else
                targetWs.Cells(startRow + 1, 9).value = amountRange.Cells(i, 1).value * -1
                targetWs.Cells(startRow + 1, 10).value = 1
                'targetWs.Cells(startRow + 1, 6).Value = ""
            End If
            '< ...
            targetWs.Cells(startRow, 10).value = 1
            
            ' > adding the UP reconcile note
            reconcileNoteRow = targetWs.Cells(startRow, 11).End(xlUp).Row
            Do Until reconcileNoteRow = 1
                targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
                reconcileNoteRow = targetWs.Cells(reconcileNoteRow, 11).End(xlUp).Row
            Loop
            ' > adding the DOWN reconcile note
            reconcileNoteRow = targetWs.Cells(startRow, 11).End(xlDown).Row
            If targetWs.Cells(reconcileNoteRow, 1).value = datesRange.Cells(i, 1).value Then
                targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
            Else
                ' reconcile that belongs to another date
            End If
            ' > adding the TOP reconcile note
            ' > @TODO what about most top reconcile??
            reconcileNoteRow = targetWs.Cells(2, 11).Row 'weird
            ' Do ?
            
            ' Loop ?
            targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
    
            
            '
        Else
            'duplicate entry found
            Debug.Print "Duplicate found"
            datesRange.Cells(i, 1).Interior.ColorIndex = 3
            amountRange.Cells(i, 1).Interior.ColorIndex = 3
        End If
    Next i
    ' To not forget. it is a error vector.
    WS.Cells(2, 1).value = ""
' Error Handlers
    Exit Sub
WRONGDATACOUNT:
    MsgBox ("wrong data count.")
    Exit Sub
WRONGPAGENAME:
    MsgBox ("wrong page name you give.")
    Exit Sub
End Sub
Private Function MakeDecimalCalculations(a As Double, b As Double) As String
    MakeDecimalCalculations = CDec(-a) / CDec(b)
End Function


'===================================================
'>> Check Duplicates
'===================================================
Private Sub CheckDuplicate_Test()
    Debug.Print CheckDuplicate("14.02.2024", CDbl(-114.99), ActiveWorkbook.Worksheets("TEBKrediKart�"))
End Sub
Private Function CheckDuplicate(chkDate As Variant, chkAmount As Double, targetWs As Worksheet) As Long
'checks duplicate for 1 and 9th columns same time. if result is zero then value is not a duplicate.
    CheckDuplicate = 0
    Dim found As Variant
    Dim found1Arr As Variant
    Dim found2Arr As Variant
    Set found = modFindAll64.FindAll(targetWs.Cells(, 1).EntireColumn, CDate(chkDate), xlValues, xlWhole)
    If found Is Nothing Then Exit Function
    found1Arr = Split(StripNonDigits(found.Address), ",")
    Set found = Nothing
    Set found = modFindAll64.FindAll(targetWs.Cells(, 9).EntireColumn, chkAmount, xlFormulas2)
    If found Is Nothing Then Exit Function
    found2Arr = Split(StripNonDigits(found.Address), ",")
    Dim result As Long: result = 0
    Dim val As Variant: val = False
    If UBound(found2Arr) = -1 Then result = 0: Exit Function
    For Each val In found2Arr
        If UBound(Filter(found1Arr, val)) <> -1 Then result = val: Exit For
    Next val
    CheckDuplicate = result
End Function
Private Function StripNonDigits(str As String) As String
    Dim result As Variant
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        .pattern = "[A-Z]"
        .Global = True
    End With
    result = oReg.Replace(str, "")
    StripNonDigits = Replace(result, "$", "")
End Function
'===================================================
'<< Check Duplicates
'===================================================

'===================================================
'>> Find Commodity Count
'===================================================
Private Function CommodityCount(str As String) As Double
    Dim result As Variant
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        .pattern = "(\d+ Pay)|(x\d+.\d+)"
    End With
    Set result = oReg.execute(str)
    If result.Count = 1 Then
        result = result(0).value
        If left(result, 1) = "x" Then
            result = CDbl(Replace(Replace(result, "x", ""), ".", ","))
        ElseIf right(result, 1) = "y" Then
            result = CDbl(Replace(Replace(result, " Pay", ""), ".", ","))
        Else
            Debug.Print "unknown transaction format"
            result = 666666
        End If
        CommodityCount = result
    Else
        CommodityCount = 666666
    End If
End Function
'===================================================
'<< Find Commodity Count
'===================================================
FİS\\2.TAŞERON_HAKEDİŞLERİ\\Z-CALISMA-HKD-2\\Berk\\Fiyat%20Güncelleme\\Elektromak\\","file:///\\\\10.20.0.13\\share\\5.�