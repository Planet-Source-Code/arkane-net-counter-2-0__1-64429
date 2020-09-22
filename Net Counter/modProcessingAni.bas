Attribute VB_Name = "modProcessingAni"
Public Sub ProcessingAnimation()

Dim count As Long

If count = 5 Then count = 1

Select Case count
Case 1
    FrmOptions.LblProcessing.Caption = "Processing"
    count = count + 1
    Exit Sub
Case 2
    FrmOptions.LblProcessing.Caption = "Processing >"
    count = count + 1
    Exit Sub
Case 3
    FrmOptions.LblProcessing.Caption = "Processing  >"
    count = count + 1
    Exit Sub
Case 4
    FrmOptions.LblProcessing.Caption = "Processing   >"
    count = count + 1
    Exit Sub
Case 5
    FrmOptions.LblProcessing.Caption = "Processing    >"
    count = count + 1
    Exit Sub
End Select

End Sub


