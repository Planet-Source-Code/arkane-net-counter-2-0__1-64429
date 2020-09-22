Attribute VB_Name = "modLogToExcel"
'Purpose: Export log data to Excel for calculations and analysis

Public Sub LogToExcel(ByVal LogPath As String, IncludeMinField As Boolean)

On Error Resume Next

Dim TextLine    As String
Dim ColCount    As Long
Dim RowCount    As Long
Dim x           As Long

'Check validity of file (before opening)
If FileLen(LogPath) = 0 Then
   MsgBox "This file appears to be empty. Cannot continue", vbInformation, App.EXEName
   Exit Sub
End If

'Open Log file
Open LogPath For Input As #1

'Check validity of file (after opening)
Line Input #1, TextLine
If LTrim(TextLine) <> "Net Counter" Then
    MsgBox "This file does not appear to be a net counter file. Cannot continue", vbInformation, App.EXEName
    Close #1
    Exit Sub
End If

'Excel
RowCount = -1
ColCount = 1

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True  '= False
objExcel.Workbooks.add

Do While Not EOF(1)
    Line Input #1, TextLine

    InfoArray = Split(TextLine, "|", -1)
    If UBound(InfoArray) <> 0 Then
        For x = 0 To UBound(InfoArray)
            objExcel.Cells(RowCount, (x + 1)).Value = InfoArray(x)
            If (IncludeMinField = True) And (x = 2) Then 'For the minutes field
                objExcel.Cells(1, (x + 5)).Value = "Time Elapsed (min)"
                ElapsedTimeArray = Split(InfoArray(x), ":", -1)
                objExcel.Cells(RowCount, (x + 5)).Value = TimeToMin(ElapsedTimeArray(0), ElapsedTimeArray(1), ElapsedTimeArray(2))
            End If
        Next
    RowCount = RowCount + 1
    End If
        
    For FormatCol = 1 To 7 'Make headings bold
        objExcel.Cells(1, FormatCol).Font.Bold = True
        objExcel.Cells(1, FormatCol).Font.Size = 12
    Next
Loop
Close #1

objExcel.Columns("A:F").AutoFit 'Autofit the columns to contents
MsgBox "Export Complete", vbInformation, App.EXEName
'objExcel.Visible = True

End Sub

Private Function TimeToMin(ByVal Hour As Long, ByVal Minutes As Long, ByVal Secs As Long) As String

Dim TimeInMin As Long

TimeInMin = (Hour * 60) + Minutes + (Secs / 60)

TimeToMin = TimeInMin

End Function
