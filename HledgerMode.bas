Attribute VB_Name = "HledgerMode"
Option Explicit

Public Sub InvokeHledgerMode()

    Application.OnKey "%+{X}", "RunHledgerCommand"
    Application.OnKey "%+{C}", "UnInvokeHledgerMode"

End Sub

Private Sub UnInvokeHledgerMode()
    
    Application.OnKey "%+{X}"
    Application.OnKey "%+{C}"

End Sub

Private Sub RunHledgerCommand()

    ActiveSheet.Cells(1, 3).EntireColumn.Resize(, ActiveSheet.columns.Count - 3).value = ""
    Dim cmdText As String
    cmdText = ActiveCell.value
    If cmdText = "" Then Exit Sub
    
    Dim sh As Object
    Set sh = CreateObject("Wscript.Shell")

    Dim shResponse As Object
    Dim shOutput As Object
        
    Dim isOutputCSV As Boolean
    If InStr(UCase(cmdText), "-O CSV") <> 0 Then
        isOutputCSV = True
    Else
        cmdText = cmdText & " -O csv --commodity-column"
        isOutputCSV = True
    End If

    Set shResponse = sh.Exec("cmd.exe /u /c chcp 65001" & "&&" & "hledger " & cmdText & "")
    Set shOutput = shResponse.StdOut

    Dim outputLine As String
    Dim outputLineSplited() As String
    
    Dim i As Integer
    Dim startLagLines As Integer
    If isOutputCSV Then
        startLagLines = 1
    Else
        startLagLines = 1
    End If
    Do While Not shOutput.AtEndOfStream
        If shOutput.line > startLagLines Then
            outputLine = shOutput.ReadLine
            outputLine = ConvertCharsToTurkish(outputLine)
            If isOutputCSV Then
                outputLineSplited = Split(outputLine, Chr(34) & "," & Chr(34))
                ActiveSheet.Cells(2 + i, 3).Resize(1, UBound(outputLineSplited) + 1) = outputLineSplited
            Else
                ActiveSheet.Cells(2 + i, 3).value = "'" & outputLine
            End If
            i = i + 1
        Else
            shOutput.ReadLine
        End If
    Loop
    If isOutputCSV Then
        ActiveSheet.UsedRange.Replace what:="""", Replacement:=""
        Dim cll As Range
        On Error Resume Next
        For Each cll In ActiveSheet.UsedRange
            If cll.value <> "" Then cll.value = cll.value * 1
        Next cll
        On Error GoTo 0
    Else
    End If

End Sub

Private Function ConvertCharsToTurkish(str As String) As String

    str = Replace(str, "Ä±", "ý")
    str = Replace(str, "Ã¶", "ö")
    str = Replace(str, "Ã§", "ç")
    str = Replace(str, "ÅŸ", "þ")
    str = Replace(str, "ÄŸ", "ð")
    str = Replace(str, "Ä°", "Ý")
    str = Replace(str, "Ã–", "Ö")
    str = Replace(str, "Ãœ", "Ü")
    ConvertCharsToTurkish = str

End Function
