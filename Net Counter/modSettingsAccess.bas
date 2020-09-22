Attribute VB_Name = "modSettingsAccess"
Option Explicit
'Purpose: Read, write settings

Public Function ReadSettings(ByVal FilePath As String) As Variant
   
Dim SettingsFileExists  As String
Dim TempStorage         As Variant
Dim SettingsArr(10)       As String
Dim XArray              As Integer

SettingsFileExists = Dir(FilePath)

XArray = 0

If SettingsFileExists <> "" Then
Open FilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, TempStorage 'Read Path into a variable.
        SettingsArr(XArray) = TempStorage
        XArray = XArray + 1
    Loop
Close #1
End If

ReadSettings = Array(SettingsArr(0), SettingsArr(1), SettingsArr(2))

End Function

Public Sub CreateSettingsFile(ByVal FilePath As String, ByVal strData As String)

Open FilePath For Output As #1
Print #1, strData
Close #1
    
End Sub

Public Sub WriteSettings(FilePath As String, strData As String)

Open FilePath For Output As #1
    Print #1, strData
Close #1
    
End Sub
