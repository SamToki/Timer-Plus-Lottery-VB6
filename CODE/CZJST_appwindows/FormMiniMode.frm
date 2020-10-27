VERSION 5.00
Begin VB.Form FormMiniMode 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Timer+Lottery"
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMiniMode.frx":0000
   LinkTopic       =   "FormMiniMode"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormMiniMode.frx":0CB2
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   4200
      Top             =   1575
   End
   Begin VB.Timer TimerMiniModeResponseMonitor 
      Interval        =   5000
      Left            =   4305
      Top             =   0
   End
   Begin VB.Timer TimerMiniModeOclockBlink 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   315
   End
   Begin VB.CommandButton CmdRestartComputer 
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3255
      MouseIcon       =   "FormMiniMode.frx":0E04
      MousePointer    =   99  'Custom
      TabIndex        =   18
      ToolTipText     =   "Restart Computer"
      Top             =   1365
      Width           =   1065
   End
   Begin VB.CommandButton CmdShutDownComputer 
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      MouseIcon       =   "FormMiniMode.frx":0F56
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Shut Down Computer"
      Top             =   1365
      Width           =   1065
   End
   Begin VB.CommandButton CmdLockCurrentUser 
      Caption         =   "Lock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      MouseIcon       =   "FormMiniMode.frx":10A8
      MousePointer    =   99  'Custom
      TabIndex        =   16
      ToolTipText     =   "Lock Current User"
      Top             =   1365
      Width           =   1065
   End
   Begin VB.CommandButton CmdRunWindowsCalculator 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MouseIcon       =   "FormMiniMode.frx":11FA
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Run Windows Calculator"
      Top             =   1365
      Width           =   1065
   End
   Begin VB.CommandButton CmdLotteryStartLottery 
      Caption         =   "START LOTTERY"
      Height          =   330
      Left            =   2205
      MouseIcon       =   "FormMiniMode.frx":134C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   945
      Width           =   2115
   End
   Begin VB.CommandButton CmdTimerReset 
      Caption         =   "RESET"
      Height          =   330
      Left            =   3255
      MouseIcon       =   "FormMiniMode.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton CmdTimerStartPauseResume 
      Caption         =   "START"
      Height          =   330
      Left            =   2205
      MouseIcon       =   "FormMiniMode.frx":15F0
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton CmdFix 
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      MouseIcon       =   "FormMiniMode.frx":1742
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Fix Mini Mode so as to prevent it from shrinking automatically"
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton CmdEXIT 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3885
      MouseIcon       =   "FormMiniMode.frx":1894
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "EXIT"
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3255
      MaskColor       =   &H000000FF&
      MouseIcon       =   "FormMiniMode.frx":19E6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Close Mini Mode and restore Main Window"
      Top             =   105
      Width           =   435
   End
   Begin VB.Timer TimerMiniModeAutoHide 
      Interval        =   100
      Left            =   2415
      Top             =   0
   End
   Begin VB.Timer TimerMiniModeDotBlink 
      Interval        =   500
      Left            =   420
      Top             =   315
   End
   Begin VB.Label LabelLotteryDisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1155
      TabIndex        =   13
      Top             =   975
      Width           =   842
   End
   Begin VB.Shape ShapeLightLottery 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   225
      Left            =   840
      Shape           =   3  'Circle
      Top             =   984
      Width           =   225
   End
   Begin VB.Label LabelLotteryTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   236
      Left            =   105
      TabIndex        =   12
      Top             =   1004
      Width           =   690
   End
   Begin VB.Label LabelTimerDisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1155
      TabIndex        =   9
      Top             =   551
      Width           =   840
   End
   Begin VB.Shape ShapeLightTimer 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      Height          =   225
      Left            =   840
      Shape           =   3  'Circle
      Top             =   571
      Width           =   225
   End
   Begin VB.Label LabelTimerTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "TIMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   236
      Left            =   105
      TabIndex        =   8
      Top             =   591
      Width           =   690
   End
   Begin VB.Label LabelClockDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00/00 (00)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   1100
      TabIndex        =   4
      Top             =   100
      Width           =   980
   End
   Begin VB.Label LabelClockSec 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   246
      Left            =   732
      TabIndex        =   3
      Top             =   120
      Width           =   280
   End
   Begin VB.Label LabelClockMin 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   246
      Left            =   431
      TabIndex        =   2
      Top             =   59
      Width           =   280
   End
   Begin VB.Label LabelClockDot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   246
      Left            =   331
      TabIndex        =   1
      Top             =   59
      Width           =   120
   End
   Begin VB.Label LabelClockHour 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   246
      Left            =   60
      TabIndex        =   0
      Top             =   59
      Width           =   280
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   2030
      Left            =   -200
      Top             =   -200
      Width           =   4650
   End
End
Attribute VB_Name = "FormMiniMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Public minimodeclockoclockblinkongoing As Boolean
Public minimodeclockoclockblinkrepeatedtimes As Integer
Public minimodeclockoclockblinkopacity As Integer
Public minimodeclockoclockblinksetopacity As Integer
Public minimoderesponsemonitorbefore As Integer
Public minimoderesponsemonitorafter As Integer

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer

    'ALWAYS FRONT (CODES FROM INTERNET)
        Dim retValue As Long
        Private Declare Function SetWindowPos Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal X As Long, ByVal Y As Long, _
            ByVal cX As Long, ByVal cY As Long, _
            ByVal wFlags As Long _
            ) As Long
            Const HWND_TOPMOST = -1
            Const SWP_SHOWWINDOW = &H40

    'HALF TRANSPARENT (CODES FROM INTERNET)
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
        Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
        '其中hwnd是透明窗体的句柄，crKey为颜色值，bAlpha是透明度，取值范围是[0,255]，dwFlags是透明方式，可以取两个值：当取值为LWA_ALPHA时，crKey参数无效，bAlpha参数有效；当取值为LWA_COLORKEY时，bAlpha参数有效而窗体中的所有颜色为crKey的地方将变为透明－－这个功能很有用：我们不必再为建立不规则形状的窗体而调用一大堆区域分析、创建、合并函数了，只需指定透明处的颜色值即可
        Private Const GWL_EXSTYLE = (-20)
        Private Const LWA_COLORKEY = &H1
        Private Const LWA_ALPHA = &H2
        Private Const ULW_COLORKEY = &H1
        Private Const ULW_ALPHA = &H2
        Private Const ULW_OPAQUE = &H4
        Private Const WS_EX_LAYERED = &H80000

        Public Function isTransparent(ByVal hWnd As Long) As Boolean
            On Error Resume Next
            Dim Msg As Long
            Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
            If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
                isTransparent = True
            Else
                isTransparent = False
            End If
            If Err Then
                isTransparent = False
            End If
        End Function
 
        Public Function MakeTransparent(ByVal hWnd As Long, ByVal Perc As Integer) As Long
            Dim Msg As Long
            On Error Resume Next
            If Perc < 0 Or Perc > 255 Then
                MakeTransparent = 1
            Else
                Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
                Msg = Msg Or WS_EX_LAYERED
                SetWindowLong hWnd, GWL_EXSTYLE, Msg
                SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
                MakeTransparent = 0
            End If
            If Err Then
                MakeTransparent = 2
            End If
        End Function
 
        Public Function MakeOpaque(ByVal hWnd As Long) As Long
            Dim Msg As Long
            On Error Resume Next
            Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
            Msg = Msg And Not WS_EX_LAYERED
            SetWindowLong hWnd, GWL_EXSTYLE, Msg
            SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
            MakeOpaque = 0
            If Err Then
                MakeOpaque = 2
            End If
        End Function

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Public Sub Form_Load()
        'ALWAYS FRONT (CODES FROM INTERNET)
        retValue = SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 1, 1, SWP_SHOWWINDOW)
        'HALF TRANSPARENT (CODES FROM INTERNET)
        MakeTransparent Me.hWnd, 255 * 0.8

        'Locate position...
        Me.Move 0, 0, 0, 0
        windowanimationtargetleft = 0
        windowanimationtargettop = 0
        Select Case FormMainWindow.minimodeclockalwaysshowdateswitch
            Case True
                windowanimationtargetwidth = 2200
            Case False
                windowanimationtargetwidth = 1050
        End Select
        windowanimationtargetheight = 400

        FormMainWindow.minimodeautohidetimeout = -1
        minimodeclockoclockblinkongoing = False
        minimodeclockoclockblinkrepeatedtimes = -1
        minimodeclockoclockblinkopacity = 255
        minimodeclockoclockblinksetopacity = 255
        minimoderesponsemonitorbefore = 101
        minimoderesponsemonitorafter = 102
    End Sub

    Public Sub MiniModeAdjustOpacity()
        MakeTransparent Me.hWnd, 255 * (FormMainWindow.minimodewindowopacity / 100)
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    Public Sub TimerMiniModeAutoHide_Timer()
        'Locate position...
        windowanimationtargetleft = 0
        windowanimationtargettop = 0
        'Auto Hide timeout...
        FormMainWindow.minimodeautohidetimeout = FormMainWindow.minimodeautohidetimeout - 1
        If FormMainWindow.minimodeautohidetimeout < 0 Then
            Select Case FormMainWindow.minimodeclockalwaysshowdateswitch
                Case True
                    windowanimationtargetwidth = 2200
                Case False
                    If (FormMainWindow.minimodetimeroverwritedateswitch = True And FormMainWindow.timerswitch = True) Then windowanimationtargetwidth = 2200 Else windowanimationtargetwidth = 1050
            End Select
            windowanimationtargetheight = 400
        Else
            windowanimationtargetwidth = 4455
            windowanimationtargetheight = 1830
        End If
        If FormMainWindow.minimodeautohidetimeout < -1000 Then FormMainWindow.minimodeautohidetimeout = -1000
    End Sub
    Public Sub TimerMiniModeDotBlink_Timer()
        Select Case FormMainWindow.minimodeclockdotblinkswitch
            Case True
                If LabelClockDot.Visible = True Then LabelClockDot.Visible = False Else LabelClockDot.Visible = True
            Case False
                LabelClockDot.Visible = True
        End Select
    End Sub
    Public Sub TimerMiniModeOclockBlink_Timer()
        minimodeclockoclockblinkongoing = True
        Select Case minimodeclockoclockblinkrepeatedtimes
            Case 111 'Ending of window blink...
                If minimodeclockoclockblinkopacity < minimodeclockoclockblinksetopacity Then minimodeclockoclockblinkopacity = minimodeclockoclockblinkopacity + 1
                If minimodeclockoclockblinkopacity > minimodeclockoclockblinksetopacity Then minimodeclockoclockblinkopacity = minimodeclockoclockblinkopacity - 1
                If minimodeclockoclockblinkopacity = minimodeclockoclockblinksetopacity Then 'End window blink...
                    TimerMiniModeOclockBlink.Enabled = False
                    minimodeclockoclockblinkongoing = False
                    minimodeclockoclockblinkrepeatedtimes = -1
                    Exit Sub
                End If
            Case 0 'Beginning of window blink...
                If minimodeclockoclockblinkopacity < 125 Then minimodeclockoclockblinkopacity = minimodeclockoclockblinkopacity + 1
                If minimodeclockoclockblinkopacity > 125 Then minimodeclockoclockblinkopacity = minimodeclockoclockblinkopacity - 1
                If minimodeclockoclockblinkopacity = 125 Then minimodeclockoclockblinkrepeatedtimes = 1
            Case Else 'Window blinking...
                'Calculate transparency...
                minimodeclockoclockblinkopacity = 125 - 125 * Sin(0.314 * (minimodeclockoclockblinkrepeatedtimes - 1))
                minimodeclockoclockblinkrepeatedtimes = minimodeclockoclockblinkrepeatedtimes + 1
        End Select
        MakeTransparent Me.hWnd, minimodeclockoclockblinkopacity
    End Sub

    Public Sub TimerMiniModeResponseMonitor_Timer()
        minimoderesponsemonitorafter = FormMiniMode.LabelClockSec
        If minimoderesponsemonitorafter = minimoderesponsemonitorbefore Then
            MsgBox "CAUTION: The Main Window of Timer+Lottery is not responding." & vbCrLf & "Fixing the problem automatically.", vbExclamation + vbOKOnly + vbDefaultButton1, "Timer+Lottery"

            FormMiniMode.Show
            FormMiniMode.windowanimationtargetleft = 0
            FormMiniMode.windowanimationtargettop = 0
            FormMiniMode.windowanimationtargetwidth = 4455
            FormMiniMode.windowanimationtargetheight = 1830
            FormMainWindow.minimodeautohidetimeout = 10
    
            If FormMainWindow.bigfloatingclockswitch = True Then Call FormMainWindow.MenuExtrasBigFloatingClock_Click

            'Call FormMainWindow...
            FormMainWindow.Show
            FormMainWindow.WindowState = 0
            'Call FormMainWindow.MenuLanguageENG_Click
            FormMiniMode.LabelLotteryDisplay.Caption = "0"
        Else
            minimoderesponsemonitorbefore = minimoderesponsemonitorafter
        End If
    End Sub

'[] COMMANDS []

    Public Sub CmdFix_Click()
        Select Case FormMainWindow.minimodeautohidesettimeout
            Case 5
                FormMainWindow.minimodeautohidesettimeout = 999
                TimerMiniModeAutoHide.Enabled = False
                CmdFix.Caption = "+"
            Case 999
                FormMainWindow.minimodeautohidesettimeout = 5
                TimerMiniModeAutoHide.Enabled = True
                CmdFix.Caption = "="
            Case Else
                MsgBox "ERROR: Mini mode auto hide set timeout is out of range." & vbCrLf & "We would appreciate it if you can send a feedback to us so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
        End Select
        Call ExpandAtClick
    End Sub
    Public Sub CmdClose_Click()
        FormMainWindow.Show
        FormMainWindow.WindowState = 0

        TimerMiniModeAutoHide.Enabled = False
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
    End Sub
    Public Sub CmdEXIT_Click()
        Call FormMainWindow.MenuEXIT_Click
    End Sub

    Public Sub CmdTimerStartPauseResume_Click()
        Call FormMainWindow.MenuTimerStartPauseResume_Click
    End Sub
    Public Sub CmdTimerReset_Click()
        Call FormMainWindow.MenuTimerReset_Click
    End Sub
    Public Sub CmdLotteryStartLottery_Click()
        Call FormMainWindow.MenuLotteryStartLottery_Click
    End Sub

    Public Sub CmdRunWindowsCalculator_Click()
        Call FormMainWindow.MenuExtrasRunWindowsCalculator_Click
    End Sub
    Public Sub CmdLockCurrentUser_Click()
        Call FormMainWindow.MenuExtrasLockCurrentUser_Click
    End Sub
    Public Sub CmdShutDownComputer_Click()
        Call FormMainWindow.MenuExtrasShutDownComputer_Click
    End Sub
    Public Sub CmdRestartComputer_Click()
        Call FormMainWindow.MenuExtrasRestartComputer_Click
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] SPECIAL []

    Public Sub ExpandAtMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If FormMainWindow.minimodeautoexpandswitch = False Then Exit Sub

        Select Case FormMainWindow.minimodeautohidesettimeout
            Case 5
                FormMainWindow.minimodeautohidetimeout = 50
                TimerMiniModeAutoHide.Enabled = True
            Case 999
                FormMainWindow.minimodeautohidetimeout = 9999
                TimerMiniModeAutoHide.Enabled = False
            Case Else
                MsgBox "ERROR: Mini mode auto hide set timeout is out of range." & vbCrLf & "We would appreciate it if you can send a feedback to us so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
        End Select
        'Locate position...
        windowanimationtargetleft = 0
        windowanimationtargettop = 0
        windowanimationtargetwidth = 4455
        windowanimationtargetheight = 1830
    End Sub
    Public Sub ExpandAtClick()
        Select Case FormMainWindow.minimodeautohidesettimeout
            Case 5
                FormMainWindow.minimodeautohidetimeout = 50
                TimerMiniModeAutoHide.Enabled = True
            Case 999
                FormMainWindow.minimodeautohidetimeout = 9999
                TimerMiniModeAutoHide.Enabled = False
            Case Else
                MsgBox "ERROR: Mini mode auto hide set timeout is out of range." & vbCrLf & "We would appreciate it if you can send a feedback to us so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
        End Select
        'Locate position...
        windowanimationtargetleft = 0
        windowanimationtargettop = 0
        windowanimationtargetwidth = 4455
        windowanimationtargetheight = 1830
    End Sub

    Public Sub LabelClockHour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call ExpandAtMouseMove(Button, Shift, X, Y)
    End Sub
    Public Sub LabelClockDot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call ExpandAtMouseMove(Button, Shift, X, Y)
    End Sub
    Public Sub LabelClockMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call ExpandAtMouseMove(Button, Shift, X, Y)
    End Sub
    Public Sub LabelClockSec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call ExpandAtMouseMove(Button, Shift, X, Y)
    End Sub
    Public Sub LabelClockDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call ExpandAtMouseMove(Button, Shift, X, Y)
    End Sub
    Public Sub LabelClockHour_Click()
        Call ExpandAtClick
    End Sub
    Public Sub LabelClockDot_Click()
        Call ExpandAtClick
    End Sub
    Public Sub LabelClockMin_Click()
        Call ExpandAtClick
    End Sub
    Public Sub LabelClockSec_Click()
        Call ExpandAtClick
    End Sub
    Public Sub LabelClockDate_Click()
        Call ExpandAtClick
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerWindowAnimation_Timer()
        If ((Me.Left = windowanimationtargetleft) And (Me.Top = windowanimationtargettop) And (Me.Width = windowanimationtargetwidth) And (Me.Height = windowanimationtargetheight)) Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If Me.Left > windowanimationtargetleft Then Me.Left = Me.Left - Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Left < windowanimationtargetleft Then Me.Left = Me.Left + Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Top > windowanimationtargettop Then Me.Top = Me.Top - Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Top < windowanimationtargettop Then Me.Top = Me.Top + Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Width > windowanimationtargetwidth Then Me.Width = Me.Width - Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Width < windowanimationtargetwidth Then Me.Width = Me.Width + Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Height > windowanimationtargetheight Then Me.Height = Me.Height - Abs(Me.Height - windowanimationtargetheight) / 4
                If Me.Height < windowanimationtargetheight Then Me.Height = Me.Height + Abs(Me.Height - windowanimationtargetheight) / 4
                If Abs(Me.Left - windowanimationtargetleft) < 10 Then Me.Left = windowanimationtargetleft
                If Abs(Me.Top - windowanimationtargettop) < 10 Then Me.Top = windowanimationtargettop
                If Abs(Me.Width - windowanimationtargetwidth) < 10 Then Me.Width = windowanimationtargetwidth
                If Abs(Me.Height - windowanimationtargetheight) < 10 Then Me.Height = windowanimationtargetheight
            Case False
                Me.Move windowanimationtargetleft, windowanimationtargettop, windowanimationtargetwidth, windowanimationtargetheight
        End Select

        If windowanimationtargetheight = 0 And Me.Height < 10 Then Me.Hide
    End Sub
