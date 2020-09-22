Attribute VB_Name = "modDetectConnection"
'---------------------------------------------------------------------------------------
' Module    : modDetectConnection
' Purpose   : Calculate the time that your internet connection is active
' Source    : From http://www.andreavb.com/ by Andrea Tincani in 'Calculate the time that your internet connection is active' article
'---------------------------------------------------------------------------------------

Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long

Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
Public Const ERROR_SUCCESS = 0

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Public CnName   As String
Public CnTime   As String

'Converts the string byte array returned by the API call into a VB String
Public Function ByteToString(ByteArray() As Byte) As String
    Dim I As Integer

    ByteToString = ""
    I = 0
    Do While ByteArray(I) <> 0
        ByteToString = ByteToString & Chr(ByteArray(I))
        I = I + 1
    Loop
End Function

'Converts a Single number into a time string
Private Function toTime(ByVal x As Single) As String
    toTime = Format(x Mod 60, "00")
    toTime = ":" & toTime
    x = x \ 60
    toTime = Format(x Mod 60, "00") & toTime
    toTime = ":" & toTime
    x = x \ 60
    toTime = Format(x, "00") & toTime
End Function

'Check the RAS Connections
Public Sub CheckRASConnections()
    Dim I                   As Long
    Dim RasConn(255)        As RasConn
    Dim structSize          As Long
    Dim ConnectionsCount    As Long
    Dim ret                 As Long
    Static LastTime         As Single
    Dim ElapsedTime         As Single

    If LastTime = 0 Then LastTime = Timer
    
    'Fills the RasConn structure with the data of all the opened RAS connections
    RasConn(0).dwSize = RAS_RASCONNSIZE
    structSize = RAS_MAXENTRYNAME * RasConn(0).dwSize
    ret = RasEnumConnections(RasConn(0), structSize, ConnectionsCount)
    ElapsedTime = Timer - LastTime
    
    If ElapsedTime < 0 Then ElapsedTime = 0
    'Each call to the Sub recalculate the elapsed time for all the active or new RAS connections
    If ret = ERROR_SUCCESS Then
        For I = 0 To ConnectionsCount - 1
            On Error GoTo NewConnection
            'Update an existing list item connection
            FrmMain.LstVBuffer.ListItems("K" & RasConn(I).hRasConn).Tag = FrmMain.LstVBuffer.ListItems("K" & RasConn(I).hRasConn).Tag + ElapsedTime
            
            FrmMain.LblElapsedTime.Caption = toTime(FrmMain.LstVBuffer.ListItems("K" & RasConn(I).hRasConn).Tag)
            FrmMain.LblConnectionName.Caption = "Online using " & ByteToString(RasConn(I).szEntryName)
            FrmMain.mnuTime.Caption = toTime(FrmMain.LstVBuffer.ListItems("K" & RasConn(I).hRasConn).Tag)
            CnName = ByteToString(RasConn(I).szEntryName)
            
            GoTo NextConnection

NewConnection:
            'Create a new list item connection
            CnTime = Format(Time, "hh:mm:ss")
            FrmMain.LstVBuffer.ListItems.add , "K" & RasConn(I).hRasConn, ByteToString(RasConn(I).szEntryName)
            FrmMain.LstVBuffer.ListItems("K" & RasConn(I).hRasConn).Tag = 0
            CnName = ByteToString(RasConn(I).szEntryName)
            
NextConnection:
        Next
    End If

LastTime = Timer
End Sub

'Tells you if a Connection as been started or terminated
'Returns the number of new connections, if the number is greater than
'zero it indicates the number of new connections started, if the number
'is negatice it indicates the number of connections terminated, zero if
'the number of RAS connections is the same

Function CheckConnections() As Integer
    Static ConnCount As Integer
    Dim RasConn(255) As RasConn
    Dim structSize As Long
    Dim ConnectionsCount As Long
    Dim ret As Long

    'Fills the RasConn structure with the data of all the opened RAS connections
    RasConn(0).dwSize = RAS_RASCONNSIZE
    structSize = RAS_MAXENTRYNAME * RasConn(0).dwSize
    ret = RasEnumConnections(RasConn(0), structSize, ConnectionsCount)
    CheckConnections = ConnectionsCount - ConnCount
    ConnCount = ConnectionsCount
End Function

