VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Net Counter"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin NetCounter.isButton CmdClose 
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FrmMain.frx":27A2
      Style           =   10
      Caption         =   "E&xit"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NetCounter.isButton CmdSentToTray 
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Icon            =   "FrmMain.frx":27BE
      Style           =   10
      Caption         =   "Send To Tray"
      IconAlign       =   3
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NetCounter.isButton CmdDiscon 
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FrmMain.frx":27DA
      Style           =   10
      Caption         =   "Disconnect"
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
   End
   Begin NetCounter.isButton CmdOptions 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FrmMain.frx":27F6
      Style           =   10
      Caption         =   "&Options"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Timer tmrShowPopup 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   1560
   End
   Begin VB.Timer TmrDisconnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1560
   End
   Begin VB.CheckBox ChkDiscon 
      Caption         =   "Enable Auto Disconnect"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox ChkMaxWin 
      Caption         =   "Maximize on disconnect if sent to tray"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Timer TmrConnectedTimer 
      Interval        =   500
      Left            =   4560
      Top             =   0
   End
   Begin MSComctlLib.ListView LstVBuffer 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Timer TmrCountDown 
      Interval        =   1000
      Left            =   5880
      Top             =   1080
   End
   Begin VB.Label LblDescript 
      Caption         =   "Sending to tray in:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label LblDisconnect 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label LblDescript 
      Caption         =   "Disconnecting in:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LblCountdown 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Click to pause"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblNow 
      Caption         =   "Current Date && Time"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label LblDescript 
      Caption         =   "Time Elapsed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label LblDescript 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblElapsedTime 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label LblConnectionName 
      Caption         =   "Offline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuTerminate 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMaximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTime 
         Caption         =   "00:00:00"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************************
'Future Enhancements (probable):
'A new 'Notes' field to be logged
'Ability to open log from within program

'Past Enhancements:
'2.0.33 - Changed all buttons to custom isButtton control from Fred.cpp (PSC)
'2.0.25 - Added disconnection notification
'2.0.21 - Added disconnection timer
'2.0.18 - Added Possibility to disconnect from within program
'*********************************************************************************

'Private Declare Function ShellExecute _
'Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
'ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Dll call to open file

Public TrayCountdown    As Double
Public DisconCountdown  As Double
Private SettingsFile    As String
Dim ConnEstablished     As Boolean
Public WithEvents m_NotificationWindow  As frmNotification 'For Disconnect notification
Attribute m_NotificationWindow.VB_VarHelpID = -1

Private Sub ChkDiscon_Click()

If (ChkDiscon.Value = 1) And (TmrDisconnect.Enabled = True) Then  'Check for auto disconnection option
    LblDisconnect.Caption = SecondsToLongTime(DisconCountdown)
End If

End Sub

Private Sub Form_Initialize()

'Check if the program is already running & inform the user if so
If App.PrevInstance Then
    MsgBox App.EXEName & " is already running!", vbInformation, App.EXEName
    End
End If

End Sub

Private Sub Form_Load()

Dim TempTimeArr         'Array
Dim SettingsData        'Array
Dim SettingsFileExists  As String
Set g_NotificationRequests = New Collection 'For Disconnect Notification

LblNow.Caption = Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm:ss")
TmrCountDown.Interval = 1000

SettingsFile = App.Path & "\Settings.set"
SettingsFileExists = Dir(SettingsFile)

If SettingsFileExists = "" Then 'Settings File does not exist
    'Create a settings file (Settings.set) as it does not exist
    CreateSettingsFile SettingsFile, "Target|" & getSpecialFolder(&H5, Me.hwnd) & vbCrLf & _
    "Tray|600" & vbCrLf & "Disconnect|3600"
    TrayCountdown = 600 '10 minutes
    DisconCountdown = 3600
Else
    'Read settings from file
    SettingsData = ReadSettings(SettingsFile) 'Store it in this array
    TempTimeArr = Split(SettingsData(1), "|")
    TrayCountdown = TempTimeArr(1)
    
    TempTimeArr = Split(SettingsData(2), "|") 'Reuse the variable
    DisconCountdown = TempTimeArr(1)
End If

LblCountdown.Caption = SecondsToLongTime(TrayCountdown)
ConnEstablished = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
            Case WM_MOUSEMOVE 'Mousemove!
            Case WM_LBUTTONDBLCLK 'Left DoubleClick
            Case WM_LBUTTONDOWN 'Left click
            Case WM_LBUTTONUP 'Left MouseUp (One click)
                ShowFromTray
            Case WM_RBUTTONDOWN 'Right MouseDown (One click)
                PopupMenu mnuFile 'Show a menu
            Case WM_RBUTTONUP 'Right MouseUp
            Case WM_RBUTTONDBLCLK 'Right DoubleClick
            End Select
    End If
        
End Sub

Private Sub CmdOptions_Click()

TmrCountDown.Enabled = False
Me.Enabled = False
FrmOptions.Show

End Sub

Private Sub CmdDiscon_Click()

TerminateRAS
If mnuTerminate.Enabled = True Then mnuTerminate.Enabled = False
If CmdDiscon.Enabled = True Then CmdDiscon.Enabled = False

End Sub

Private Sub CmdSentToTray_Click()

    'Procedures to add an icon to the system tray
    'Add an icon.  This procedure uses the icon specified in
    'the Icon property of the form. This can be modified as desired.

    Dim I As Integer
    Dim nid As NOTIFYICONDATA
    
   
    nid = setNotifyIconData(Me.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, Me.Icon, App.EXEName)
    I = Shell_NotifyIconA(NIM_ADD, nid)
    
    TmrCountDown.Enabled = False
    Me.WindowState = vbMinimized 'Minimize the current window
    Me.Visible = False 'Hide it
    

End Sub

Private Sub CmdClose_Click()

Unload Me

End Sub

Private Sub mnuExit_Click() 'Forms part of menu that appears when icon in tray is right-clicked

Unload Me

End Sub

Private Sub mnuMaximize_Click()

ShowFromTray

End Sub

Private Sub mnuTerminate_Click()

TerminateRAS 'Disconnect
If mnuTerminate.Enabled = True Then mnuTerminate.Enabled = False
If CmdDiscon.Enabled = True Then CmdDiscon.Enabled = False

End Sub

Private Sub LblCountdown_Click()

If TmrCountDown.Enabled = True Then
    TmrCountDown.Enabled = False
    LblCountdown.ToolTipText = "Click to resume"
Else
    TmrCountDown.Enabled = True
    LblCountdown.ToolTipText = "Click to pause"
End If

End Sub

Private Sub TmrConnectedTimer_Timer() 'This timer periodically checks if connection has been established
'Do NOT disable this timer or the program will not function as intended

Dim NumOfConnections    As Integer
Dim LogFile             As String
Dim SettingsFileExists  As String
Dim SettingsData
Dim DestFldr

NumOfConnections = CheckConnections

CheckRASConnections

If ConnEstablished = True Then
    If ChkDiscon.Value = 1 Then 'Check for auto disconnection option
        TmrDisconnect.Enabled = True
    ElseIf ChkDiscon.Value = 0 Then
        TmrDisconnect.Enabled = False
    End If
End If

'Connections monitoring
If NumOfConnections > 0 Then
    'MsgBox NumOfConnections & " New Connection Started"
    If mnuTerminate.Enabled = False Then mnuTerminate.Enabled = True
    If CmdDiscon.Enabled = False Then CmdDiscon.Enabled = True
    ConnEstablished = True
ElseIf NumOfConnections < 0 Then
    'MsgBox Abs(NumOfConnections) & " Connection Terminated"
    LblConnectionName.Caption = "Offline"
    
    If mnuTerminate.Enabled = True Then mnuTerminate.Enabled = False 'In case user disconnects using Windows interface
    If CmdDiscon.Enabled = True Then CmdDiscon.Enabled = False

    'Read settings file for where to store log
    SettingsFileExists = Dir(SettingsFile)

    If SettingsFileExists <> "" Then 'Will continue only if file exists
        SettingsData = ReadSettings(SettingsFile)
        DestFldr = Split(SettingsData(0), "|")
        LogFile = DestFldr(1) & "\Net Time Counter " & Year(Date) & ".log"
    End If
    
    SaveTime LogFile, modDetectConnection.CnTime, LblElapsedTime.Caption, Format(Time(), "hh:mm:ss"), modDetectConnection.CnName
    
    If Not Visible Then
        If ChkMaxWin.Value = 1 Then ShowFromTray
    End If
    
    If ChkDiscon.Value = 1 Then TmrDisconnect.Enabled = False
    ConnEstablished = False
End If

LblNow.Caption = Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm:ss")

End Sub

Private Sub TmrCountDown_Timer() 'Send to tray countdown timer

Dim I     As Integer
Dim nid   As NOTIFYICONDATA

TrayCountdown = Int(TrayCountdown) - 1 'Int has been used so as to cater for decimal
                                       'values (Abnormal data here)

If TrayCountdown >= 0 Then LblCountdown.Caption = SecondsToLongTime(TrayCountdown)
If TrayCountdown = 60 Then LblCountdown.FontBold = True '1 minute (60 seconds) left. Make display bold
If TrayCountdown = 0 Then 'Countdown terminates
    nid = setNotifyIconData(Me.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, Me.Icon, App.EXEName)
    I = Shell_NotifyIconA(NIM_ADD, nid)
    Me.WindowState = vbMinimized
    Me.Visible = False
    TmrCountDown.Enabled = False
End If

End Sub

Private Sub ShowFromTray()

Visible = True
WindowState = vbNormal
Dim hProcess As Long
GetWindowThreadProcessId hwnd, hProcess
AppActivate hProcess

'Activate countdown timers & displays
Dim SettingsData
Dim TempTimeVar

LblCountdown.FontBold = False
SettingsData = ReadSettings(SettingsFile)
TempTimeVar = Split(SettingsData(1), "|", -1)
TrayCountdown = TempTimeVar(1) 'Retrieve timer settings

LblCountdown.Caption = SecondsToLongTime(TrayCountdown)
TmrCountDown.Enabled = True

'Remove icon from tray.

Dim I   As Integer
Dim nid As NOTIFYICONDATA

nid = setNotifyIconData(Me.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, Me.Icon, "")
I = Shell_NotifyIconA(NIM_DELETE, nid)
                
End Sub

Private Function SaveTime(OutputFilePath As String, strStartTime As String, ElapsedTime As String, Endtime As String, ConnectionName As String)

Dim OutputFileExist  As String
Dim TextLine         As String
Dim OutputFile       As String

Dim InfoArray()      As String
Dim DateArray()      As String

Dim FileExisted      As Boolean

OutputFile = OutputFilePath

OutputFileExist = Dir(OutputFile) 'Check if output file already exists

If OutputFileExist <> "" Then 'Will continue only if file exists
    Open OutputFile For Input As #1
        Do While Not EOF(1)   'Loop until end of file.
            Line Input #1, TextLine    'Read last line into a variable.
        Loop
        InfoArray = Split(TextLine, "|", 1)
        DateArray = Split(InfoArray(0), "/")
    Close #1
    FileExisted = True 'Yes, file existed
Else
    FileExisted = False 'No, file did not exist previously
End If

    Open OutputFile For Append As #2
    
        If OutputFileExist = "" Then 'This wil be output only if the output file did not exist
            Print #2, Tab(30); "Net Counter"
            Print #2,
            Print #2, "Log creation date: " & Format(Date, "Long date")
            Print #2,
            Print #2, "---------------------------------------------------------------------------------------"
            Print #2, Spc(4); "Date"; Spc(4); "|"; Spc(1); "Start Time"; Spc(1); "|"; Spc(1); "Time elapsed"; Spc(1); "|"; Spc(1); "End Time"; Spc(1); "|"; Spc(3); "Day"; Spc(3); "|"; Spc(1); "Connection"
            Print #2, "---------------------------------------------------------------------------------------"
        End If
        
        If FileExisted = True Then
            If Month(Date) > Val(DateArray(1)) Then 'If present month is greater than the stored month
                Print #2,
                Print #2, "-------------------------- " & MonthName(Month(Date), False) & " " & Year(Date) & " -----------------------------"
                Print #2,
            End If
        End If
        Print #2, Spc(1); Format(Date, "dd/mm/yyyy"); Spc(1); "|"; Spc(2); strStartTime; Spc(2); "|"; Spc(3); ElapsedTime; Spc(3); "|"; Spc(1); Endtime; Spc(1); "|"; Spc(1); WeekdayName(Weekday(Date, vbMonday), False, vbMonday); Spc(3); "|"; Spc(1); ConnectionName
    Close #2

End Function

Private Sub TmrDisconnect_Timer() 'Timer that countdowns till disconnection

Static m_ComputerCount As Long
Dim lNotificationRequest      As cNotificationRequest

DisconCountdown = Int(DisconCountdown) - 1

If DisconCountdown = 60 Then 'Show Notification on 60 seconds
    tmrShowPopup.Enabled = True
    m_ComputerCount = m_ComputerCount + 1
    Call mNotificationSystem.RequestUserNotification("NOTIFY: " & "Host " & CStr(m_ComputerCount), App.EXEName & ":" & vbCrLf & " Disconnection Warning", "Disconnecting in: " & SecondsToLongTime(DisconCountdown), True, True)
    'tmrShowPopup.Enabled = False
End If

If DisconCountdown >= 0 Then LblDisconnect.Caption = SecondsToLongTime(DisconCountdown)
If DisconCountdown = 0 Then TerminateRAS 'Terminate connection

End Sub

Private Sub tmrShowPopup_Timer()

Dim lNotificationRequest      As cNotificationRequest

    ' Check if we have some requests and make sure we are'nt showing a notification already.
    If (g_NotificationRequests.Count > 0) And (m_NotificationWindow Is Nothing) Then  '(Not IsFormLoaded("frmUserPopup")) Then
        ' Get the first Notification Request from the Collection.
        Set lNotificationRequest = g_NotificationRequests.Item(1)
        
        ' Setup and Show the notification request.
        Set m_NotificationWindow = New frmNotification
        Call m_NotificationWindow.ShowNotification(lNotificationRequest)
        
        ' Remove the Request from the Collection.
        g_NotificationRequests.Remove 1
    End If
    
    Set lNotificationRequest = Nothing

End Sub

Private Sub m_NotificationWindow_Finished()
    ' Set the Notification variable to nothing, to indicate that we have
    ' finished using it.
    Set m_NotificationWindow = Nothing
End Sub
