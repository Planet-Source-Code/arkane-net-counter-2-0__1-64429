VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin NetCounter.isButton CmdOk 
      Height          =   585
      Left            =   5040
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   820
      Icon            =   "FrmOptions.frx":27A2
      Style           =   10
      Caption         =   "OK"
      IconAlign       =   0
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin NetCounter.isButton CmdBrowse 
      Height          =   660
      Left            =   5040
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1164
      Icon            =   "FrmOptions.frx":27BE
      Style           =   10
      Caption         =   "Change Location"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Frame FrOpt 
      Caption         =   "Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6015
      Begin VB.ComboBox CboDisconTimes 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox CombTime 
         Height          =   315
         ItemData        =   "FrmOptions.frx":27DA
         Left            =   1560
         List            =   "FrmOptions.frx":27DC
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblDescript 
         Caption         =   "minutes if enabled"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label LblDescript 
         Caption         =   "Disconnect in:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label LblDescript 
         Caption         =   "Minimize to tray in:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label LblPostfix 
         Caption         =   "minutes"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FrExcelOpt 
      Caption         =   "Excel Options"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4575
      Begin NetCounter.isButton CmdExprtExcel 
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Icon            =   "FrmOptions.frx":27DE
         Style           =   10
         Caption         =   "Export log to Excel"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Include a minutes field"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.TextBox TxtLocation 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label LblDescript 
      Caption         =   "Save log file to:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SettingsFile    As String
Private SettingData 'Will be used as an array. Accessible anywhere within this module (form)

Private Sub CboDisconTimes_Click()

Dim SettingsFileExists As String

SettingsFileExists = Dir(SettingsFile)

If SettingsFileExists <> "" Then 'Will continue only if file exists
    WriteSettings SettingsFile, SettingData(0) & vbCrLf & SettingData(1) & vbCrLf & "Disconnect|" & (CboDisconTimes.Text * 60)
    FrmMain.DisconCountdown = (CboDisconTimes.Text * 60)
    If FrmMain.TmrDisconnect.Enabled = True Then FrmMain.LblDisconnect.Caption = SecondsToLongTime(FrmMain.DisconCountdown)
End If

End Sub

Private Sub CmdBrowse_Click()

Dim PreviousSetting As String
Dim sFile           As String

PreviousSetting = TxtLocation.Text
sFile = BrowseForFolder("c:\", Me.hwnd, "Browse for Folder")

If sFile = "" Then
    TxtLocation.Text = PreviousSetting
Else
    TxtLocation.Text = sFile
    WriteSettings SettingsFile, "Target|" & sFile & vbCrLf & _
    SettingData(1) & vbCrLf & SettingData(2)
End If

End Sub

Private Sub CmdExprtExcel_Click()

Dim LogLocation As String
Dim IncMinField As Boolean

cmnDlg.Dialogtitle = "Open a " & App.EXEName & " log file"
cmnDlg.Filename = ""
cmnDlg.Initdir = TxtLocation.Text 'Initial directory
cmnDlg.Filter = "Log Files (*.log)|*.log|Text file (*.txt)|*.txt"
cmnDlg.Flags = 5 'to remove the 'Readonly' checkbox
ShowOpen

If ChkOpt(0).Value = 0 Then 'Check if user wants to include a minute field
    IncMinField = False
Else
    IncMinField = True
End If

If Len(cmnDlg.Filename) = 0 Then
    Exit Sub 'Use chooses Cancel
Else
    LogLocation = cmnDlg.Filename
    LogToExcel LogLocation, IncMinField
End If

End Sub

Private Sub CmdOK_Click()

Unload Me

End Sub

Private Sub CombTime_Click()

Dim SettingsFileExists As String

SettingsFileExists = Dir(SettingsFile)

If SettingsFileExists <> "" Then 'Will continue only if file exists
    WriteSettings SettingsFile, SettingData(0) & vbCrLf & "Tray|" & (CombTime.Text * 60) & vbCrLf & SettingData(2)
    FrmMain.TrayCountdown = (CombTime.Text * 60)
    FrmMain.LblCountdown.Caption = SecondsToLongTime(FrmMain.TrayCountdown)
End If

End Sub

Private Sub Form_Load()


Dim SettingsFileExists  As String
Dim Alpha               As Integer
Dim TimeValue           As Long
Dim DestFldr
Dim TimeSetting

'Read settings from a settings file (Settings.set) if it exists
SettingsFile = App.Path & "\Settings.set" 'Store in variable path where file should normally be
SettingsFileExists = Dir(SettingsFile)

If SettingsFileExists <> "" Then 'Will continue only if file exists
    'Stores the settings in an array in memory
    SettingData = ReadSettings(SettingsFile) 'Returns an array

    DestFldr = Split(SettingData(0), "|") 'Processing first line of settings file
    TxtLocation.Text = DestFldr(1)
Else
    CreateSettingsFile SettingsFile, "Target|" & getSpecialFolder(&H5, Me.hwnd) & vbCrLf & _
    "Tray|600" & vbCrLf & "Disconnect|3600" 'Create a settings file with default settings
    TxtLocation.Text = getSpecialFolder(&H5, Me.hwnd)
End If

'Combobox CombTime
TimeValue = 5
For Alpha = 1 To 12
    CombTime.AddItem TimeValue
    TimeValue = TimeValue + 5
Next

TimeSetting = Split(SettingData(1), "|", -1)
TimeValue = TimeSetting(1) 'Reuse TimeValue Variable
CombTime.ListIndex = ((TimeValue \ 300) - 1) 'Dynamically calculate what item to display

'Combobox CboDisconTimes
TimeValue = 5
For Alpha = 1 To 24
    CboDisconTimes.AddItem TimeValue
    TimeValue = TimeValue + 5
Next

TimeSetting = Split(SettingData(2), "|", -1)
TimeValue = TimeSetting(1) 'Reuse TimeValue Variable
CboDisconTimes.ListIndex = ((TimeValue \ 300) - 1) 'Dynamically calculate what item to display


End Sub

Private Sub Form_Unload(Cancel As Integer)

FrmMain.Enabled = True
FrmMain.TmrCountDown.Enabled = True
FrmMain.LblCountdown.ToolTipText = "Click to pause"

Unload Me

End Sub
