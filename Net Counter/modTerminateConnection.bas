Attribute VB_Name = "modTerminateConnection"
'---------------------------------------------------------------------------------------
' Module        : modTerminateConnection
' Purpose       : Terminate an internet connection
' Dependencies  : Module is independent
' Source        : From http://www.andreavb.com/ by Andrea Tincani in 'Terminate all the active RAS Connections' article
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

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

'Terminate all the RAS Connections
Sub TerminateRAS()
    Dim I As Long
    Dim RasConn(255) As RasConn
    Dim structSize As Long
    Dim ConnectionsCount As Long
    Dim ret As Long

    'Fills the RasConn structure with the data of all the opened RAS connections
    RasConn(0).dwSize = RAS_RASCONNSIZE
    structSize = RAS_MAXENTRYNAME * RasConn(0).dwSize
    ret = RasEnumConnections(RasConn(0), structSize, ConnectionsCount)
    'hangup all the RAS connections
    If ret = ERROR_SUCCESS Then
        For I = 0 To ConnectionsCount - 1
            ret = RasHangUp(RasConn(I).hRasConn)
        Next
    End If
End Sub

