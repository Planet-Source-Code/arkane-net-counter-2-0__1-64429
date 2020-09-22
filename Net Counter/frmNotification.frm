VERSION 5.00
Begin VB.Form frmNotification 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   75
   ClientWidth     =   5445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   Begin NetCounter.isButton CmdOK 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Icon            =   "frmNotification.frx":0000
      Style           =   4
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
   Begin NetCounter.isButton CmdDelayDiscon 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Icon            =   "frmNotification.frx":001C
      Style           =   4
      Caption         =   "Delay for 5 minutes"
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
   Begin NetCounter.isButton CmdDisable 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Icon            =   "frmNotification.frx":0038
      Style           =   4
      Caption         =   "Disable Auto Disconnect"
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
   Begin NetCounter.isButton CmdDisconnect 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Icon            =   "frmNotification.frx":0054
      Style           =   4
      Caption         =   "Disconnect Now"
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
   Begin VB.Timer tmrCtrlDiscon 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   0
   End
   Begin VB.Timer tmrPopupController 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2760
      Top             =   0
   End
   Begin VB.Shape shpBorder 
      Height          =   255
      Left            =   1560
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   960
   End
   Begin VB.Image imgNotificationIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmNotification.frx":0070
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Module     : frmNotification
'Description:
'Version    : V2.00 31/10/2005 10:17
'Release    : VB6
'Copyright  :
'Author     : Chris.Nillissen
'----------------------------------------------------------------------------------
'V2.00    31/10/2005 Original version
'
'----------------------------------------------------------------------------------

Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA    As Long = 48

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    bottom  As Long
End Type


Public Enum eStatusType
    StatusShow = 0
    StatusHide = 1
End Enum

Public Event Clicked(ByRef Key As String)
Public Event Finished()

Private m_Status                As eStatusType
Private m_FormOpenHeight        As Long
Private m_FormBottomPosition    As Long
Private m_FormRightPosition     As Long
Public m_OpenInterval          As Long

Private m_NotificationRequest   As cNotificationRequest

Private Sub CmdDelayDiscon_Click()

FrmMain.DisconCountdown = 300
Unload Me

End Sub

Private Sub CmdDisable_Click()

FrmMain.ChkDiscon.Value = 0
Unload Me

End Sub

Private Sub CmdDisconnect_Click()

TerminateRAS
If FrmMain.mnuTerminate.Enabled = True Then FrmMain.mnuTerminate.Enabled = False
If FrmMain.CmdDiscon.Enabled = True Then FrmMain.CmdDiscon.Enabled = False
Unload Me

End Sub

Private Sub CmdOK_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim lDesktopArea        As RECT
    
    ' Set default values.
    m_OpenInterval = FrmMain.DisconCountdown * 40
    
    ' Set the Window as top most window.
    Call mNotificationSystem.SetWindowTopMost(Me.hwnd)
    
    ' Get desktop area not taken up by the taskbar.
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, lDesktopArea, 0&)
    m_FormOpenHeight = Me.Height
    m_FormBottomPosition = (lDesktopArea.bottom * Screen.TwipsPerPixelY)
    m_FormRightPosition = (lDesktopArea.Right * Screen.TwipsPerPixelX)

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Disable the Popup Controller Timer.
    tmrPopupController.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Finished
End Sub

Public Property Get NotificationRequest() As cNotificationRequest
    Set NotificationRequest = m_NotificationRequest
End Property
Public Property Let NotificationRequest(ByVal vNewValue As cNotificationRequest)
    m_NotificationRequest = vNewValue
End Property

Private Sub tmrCtrlDiscon_Timer()
    Me.lblDescription.Caption = "Disconnecting in: " & SecondsToLongTime(FrmMain.DisconCountdown)
End Sub

Private Sub tmrPopupController_Timer()

    Select Case m_Status
    Case StatusShow
        tmrCtrlDiscon.Enabled = True
        Me.Move Me.Left, m_FormBottomPosition - Me.Height, Me.Width, Me.Height + 100
        If (Me.Height >= m_FormOpenHeight) Then
            Me.Height = m_FormOpenHeight
            m_Status = StatusHide
            tmrPopupController.Interval = m_OpenInterval
            Exit Sub
        End If
    Case StatusHide
        tmrCtrlDiscon.Enabled = False
        tmrPopupController.Interval = 15
        Me.Move Me.Left, m_FormBottomPosition - Me.Height, Me.Width, Me.Height - 20
        If (Me.Height < 20) Then Unload Me
    End Select

End Sub

Public Sub ShowNotification(ByVal NotificationRequest As cNotificationRequest)
    ' Store a copy of the Notification Request.
    Set m_NotificationRequest = NotificationRequest
    
    ' Setup the Window with the Notification Request settings.
    Call SetupNotification(NotificationRequest)

    ' Set starting position, size and show the window.
    Me.Move m_FormRightPosition - (Me.Width + 100), m_FormBottomPosition - 10, Me.Width, 10
    Me.Show: DoEvents
    
    ' Start showing the form starting at top of task bar.
    m_Status = StatusShow
    tmrPopupController.Enabled = True
    
    ' Play the associated wave file.
    Call mNotificationSystem.PlayWaveSoundFile(NotificationRequest.SoundFileLocation)
End Sub
Public Sub UpdateNotification(ByVal NotificationRequest As cNotificationRequest)
    ' Store a copy of the Notification Request.
    Set m_NotificationRequest = NotificationRequest
    
    ' Setup the Window with the Notification Request settings.
    Call SetupNotification(NotificationRequest)
    
    ' Start showing the form starting at top of task bar.
    m_Status = StatusShow
    tmrPopupController.Enabled = True
End Sub

Private Sub SetupNotification(ByRef NotificationRequest As cNotificationRequest)
    ' Setup the Forms Controls.
    lblTitle.Caption = NotificationRequest.Title
    lblDescription.Caption = NotificationRequest.Description
    If (Not NotificationRequest.Icon Is Nothing) Then Set imgNotificationIcon = NotificationRequest.Icon
    
    ' Size the window to fit the description.
    Me.Width = (lblDescription.Left + lblDescription.Width + 10) * Screen.TwipsPerPixelX
    Me.Left = m_FormRightPosition - (Me.Width + 100)
    
    ' Position any controls on the form.
    shpBorder.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
