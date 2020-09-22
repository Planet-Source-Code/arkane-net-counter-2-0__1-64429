Attribute VB_Name = "modTime"
'Dependencies: Module is independent

Public Function SecondsToLongTime(ByVal DblTimeInSec As Double) As String

'Input is total time in seconds. Output is in form hh:mm:ss
'Function to convert seconds into HH:MM:SS format. Is reusable

Dim hh, mm, ss
    
    hh = Int(DblTimeInSec / 3600) 'Calculate hours spent
    mm = Int(DblTimeInSec / 60) - (hh * 60) 'Calculate minutes spent
    ss = Int(DblTimeInSec Mod 60) 'Calculate seconds spent
        
    SecondsToLongTime = Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(ss, "00")
    
End Function
