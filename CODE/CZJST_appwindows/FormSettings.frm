VERSION 5.00
Begin VB.Form FormSettings 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "Timer+Lottery"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
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
   Icon            =   "FormSettings.frx":0000
   LinkTopic       =   "FormSettings"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormSettings.frx":0CB2
   MousePointer    =   99  'Custom
   ScaleHeight     =   6945
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   12600
      Top             =   6615
   End
   Begin VB.Frame FrameSounds 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Sounds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   6615
      TabIndex        =   32
      Top             =   5145
      Width           =   6105
      Begin VB.CheckBox CheckboxSoundsInterfaceSounds 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Interface Sounds"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":0E04
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   840
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxSoundsLotteryTone 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Lottery tone"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":0F56
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   420
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CheckboxSoundsTimerTone 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Timer tone"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":10A8
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   420
         Value           =   1  'Checked
         Width           =   2745
      End
   End
   Begin VB.Frame FrameDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   210
      TabIndex        =   28
      Top             =   5145
      Width           =   6105
      Begin VB.CheckBox CheckboxDisplayUseOldScrollAnimationInLotteryWindow 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Use old scroll animation in Lottery Window"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":11FA
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   840
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxDisplayWindowAnimations 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Window animations"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":134C
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   420
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CheckboxDisplayLightBulbIndicators 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Show light bulb indicators"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":149E
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   420
         Value           =   1  'Checked
         Width           =   2745
      End
   End
   Begin VB.Frame FrameMiniMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Mini Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4005
      Left            =   6615
      TabIndex        =   16
      Top             =   945
      Width           =   6105
      Begin VB.CheckBox CheckboxMiniModeAlwaysShowDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Always show date"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":15F0
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   2520
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxMiniMode24HrFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "24-Hour Format"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1742
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   3360
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxMiniModeShowSeconds 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Show seconds"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormSettings.frx":1894
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CheckboxMiniModeBlinkDots 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Blink dots on the clock"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":19E6
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   2100
         Width           =   2745
      End
      Begin VB.CheckBox CheckboxMiniModeBlinkWhenMinuteHitsZero 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Blink the window when minute hits zero"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1B38
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   1680
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxMiniModeTimerOverwriteDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Display the timer instead of the date when it is running"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1C8A
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   2940
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.HScrollBar HScrollMiniModeOpacity 
         Height          =   330
         LargeChange     =   20
         Left            =   2520
         Max             =   100
         Min             =   20
         MouseIcon       =   "FormSettings.frx":1DDC
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   480
         Value           =   80
         Width           =   3375
      End
      Begin VB.CheckBox CheckboxMiniModeDarkTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Dark Theme in Mini Mode Window"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":1F2E
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   840
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxMiniModeAutoExpand 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Automatically expand when mouse moves over"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":2080
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1260
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.Label LabelMiniModeOpacityIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "80%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   450
         Width           =   705
      End
      Begin VB.Label LabelMiniModeOpacity 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Opacity:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   510
         Width           =   1230
      End
   End
   Begin VB.Frame FrameBigFloatingClock 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Big Floating Clock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2325
      Left            =   210
      TabIndex        =   5
      Top             =   2625
      Width           =   6105
      Begin VB.CommandButton CmdBigFloatingClockPosition11 
         Caption         =   "¨K"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5565
         MouseIcon       =   "FormSettings.frx":21D2
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1785
         Width           =   330
      End
      Begin VB.CommandButton CmdBigFloatingClockPosition01 
         Caption         =   "¨L"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         MouseIcon       =   "FormSettings.frx":2324
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1785
         Width           =   330
      End
      Begin VB.CommandButton CmdBigFloatingClockPosition10 
         Caption         =   "¨J"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5565
         MouseIcon       =   "FormSettings.frx":2476
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1470
         Width           =   330
      End
      Begin VB.CommandButton CmdBigFloatingClockPosition00 
         Caption         =   "¨I"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         MouseIcon       =   "FormSettings.frx":25C8
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1470
         Width           =   330
      End
      Begin VB.VScrollBar VScrollBigFloatingClockPositionY 
         Height          =   1065
         LargeChange     =   2000
         Left            =   5565
         Max             =   1000
         MouseIcon       =   "FormSettings.frx":271A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   315
         Width           =   330
      End
      Begin VB.HScrollBar HScrollBigFloatingClockPositionX 
         Height          =   330
         LargeChange     =   2000
         Left            =   4095
         Max             =   1000
         MouseIcon       =   "FormSettings.frx":286C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1785
         Width           =   1065
      End
      Begin VB.CheckBox CheckboxBigFloatingClock24HrFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "24-Hour Format"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":29BE
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3690
      End
      Begin VB.CheckBox CheckboxBigFloatingClockBlinkWhenMinuteHitsZero 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Blink the clock when minute hits zero"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":2B10
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1260
         Value           =   1  'Checked
         Width           =   4950
      End
      Begin VB.CheckBox CheckboxBigFloatingClockShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Shadow (recommended)"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":2C62
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   420
         Value           =   1  'Checked
         Width           =   4950
      End
      Begin VB.CheckBox CheckboxBigFloatingClockAutoHide 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Automatically hide when mouse moves over"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":2DB4
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   4950
      End
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11235
      MouseIcon       =   "FormSettings.frx":2F06
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Frame FrameLottery 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Lottery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   210
      TabIndex        =   2
      Top             =   945
      Width           =   6105
      Begin VB.CheckBox CheckboxLotteryDarkTheme 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Dark Theme in Lottery Window"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":3058
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   840
         Width           =   5685
      End
      Begin VB.CheckBox CheckboxLotteryPreventSameNumbersInHistory 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Prevent same numbers in history (recommended)"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormSettings.frx":31AA
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   420
         Value           =   1  'Checked
         Width           =   5685
      End
   End
   Begin VB.Label LabelSettingsTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SETTINGS"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   315
      TabIndex        =   0
      Top             =   210
      Width           =   10515
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   6945
      Left            =   0
      Top             =   0
      Width           =   12930
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'Close button...
    Public Sub CmdClose_Click()
        windowanimationtargetleft = (Screen.Width / 2)
        windowanimationtargettop = (Screen.Height / 2)
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
    End Sub

    'Lottery settings...
    Public Sub CheckboxLotteryPreventSameNumbersInHistory_Click()
        Select Case FormMainWindow.lotterypreventrepeatswitch
            Case True
                FormMainWindow.lotterypreventrepeatswitch = False
                CheckboxLotteryPreventSameNumbersInHistory.Value = 0
            Case False
                FormMainWindow.lotterypreventrepeatswitch = True
                CheckboxLotteryPreventSameNumbersInHistory.Value = 1
        End Select
    End Sub
    Public Sub CheckboxLotteryDarkTheme_Click()
        Select Case FormMainWindow.lotterywindowdarkthemeswitch
            Case True
                FormMainWindow.lotterywindowdarkthemeswitch = False
                CheckboxLotteryDarkTheme.Value = 0
                FormLottery.BackColor = &HE0E0E0
                FormLottery.PictureboxScroll.BackColor = &HE0E0E0
                FormLottery.LabelHintText.ForeColor = &HFF6600
                FormLottery.LabelScrollText.ForeColor = &H0
            Case False
                FormMainWindow.lotterywindowdarkthemeswitch = True
                CheckboxLotteryDarkTheme.Value = 1
                FormLottery.BackColor = &H0
                FormLottery.PictureboxScroll.BackColor = &H0
                FormLottery.LabelHintText.ForeColor = &HFFAA00
                FormLottery.LabelScrollText.ForeColor = &HFFFFFF
        End Select
    End Sub

    'Big Floating Clock settings...
    Public Sub CheckboxBigFloatingClockShadow_Click()
        Select Case FormMainWindow.bigfloatingclockshadowswitch
            Case True
                FormMainWindow.bigfloatingclockshadowswitch = False
                CheckboxBigFloatingClockShadow.Value = 0
                FormBigFloatingClock.LabelHourShadow.Visible = False
                FormBigFloatingClock.LabelMinShadow.Visible = False
                FormBigFloatingClock.LabelDotShadow.Visible = False
            Case False
                FormMainWindow.bigfloatingclockshadowswitch = True
                CheckboxBigFloatingClockShadow.Value = 1
                FormBigFloatingClock.LabelHourShadow.Visible = True
                FormBigFloatingClock.LabelMinShadow.Visible = True
                FormBigFloatingClock.LabelDotShadow.Visible = True
        End Select
    End Sub
    Public Sub CheckboxBigFloatingClockAutoHide_Click()
        Select Case FormMainWindow.bigfloatingclockautohideswitch
            Case True
                FormMainWindow.bigfloatingclockautohideswitch = False
                CheckboxBigFloatingClockAutoHide.Value = 0
            Case False
                FormMainWindow.bigfloatingclockautohideswitch = True
                CheckboxBigFloatingClockAutoHide.Value = 1
        End Select
    End Sub
    Public Sub CheckboxBigFloatingClockBlinkWhenMinuteHitsZero_Click()
        Select Case FormMainWindow.bigfloatingclockoclockblinkswitch
            Case True
                FormMainWindow.bigfloatingclockoclockblinkswitch = False
                CheckboxBigFloatingClockBlinkWhenMinuteHitsZero.Value = 0
            Case False
                FormMainWindow.bigfloatingclockoclockblinkswitch = True
                CheckboxBigFloatingClockBlinkWhenMinuteHitsZero.Value = 1
        End Select
    End Sub
    Public Sub CheckboxBigFloatingClock24HrFormat_Click()
        Select Case FormMainWindow.bigfloatingclock24hrformatswitch
            Case True
                FormMainWindow.bigfloatingclock24hrformatswitch = False
                CheckboxBigFloatingClock24HrFormat.Value = 0
            Case False
                FormMainWindow.bigfloatingclock24hrformatswitch = True
                CheckboxBigFloatingClock24HrFormat.Value = 1
        End Select
        'Update immediately...
        Call FormBigFloatingClock.TimerBigFloatingClockClock_Timer
    End Sub
    Public Sub HScrollBigFloatingClockPositionX_Change()
        If FormMainWindow.bigfloatingclockswitch = False Then Call FormMainWindow.MenuExtrasBigFloatingClock_Click  'Force show effect...
        FormBigFloatingClock.windowanimationtargetleft = (HScrollBigFloatingClockPositionX.Value / 1000) * (Screen.Width - 3060)
        FormBigFloatingClock.windowanimationtargettop = (VScrollBigFloatingClockPositionY.Value / 1000) * (Screen.Height - 1485)
    End Sub
    Public Sub HScrollBigFloatingClockPositionX_Scroll()
        Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub VScrollBigFloatingClockPositionY_Change()
        Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub VScrollBigFloatingClockPositionY_Scroll()
        Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub CmdBigFloatingClockPosition00_Click()
        HScrollBigFloatingClockPositionX.Value = 0
        VScrollBigFloatingClockPositionY.Value = 0
        'Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub CmdBigFloatingClockPosition10_Click()
        HScrollBigFloatingClockPositionX.Value = 1000
        VScrollBigFloatingClockPositionY.Value = 0
        'Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub CmdBigFloatingClockPosition01_Click()
        HScrollBigFloatingClockPositionX.Value = 0
        VScrollBigFloatingClockPositionY.Value = 1000
        'Call HScrollBigFloatingClockPositionX_Change
    End Sub
    Public Sub CmdBigFloatingClockPosition11_Click()
        HScrollBigFloatingClockPositionX.Value = 1000
        VScrollBigFloatingClockPositionY.Value = 1000
        'Call HScrollBigFloatingClockPositionX_Change
    End Sub

    'Mini Mode settings...
    Public Sub HScrollMiniModeOpacity_Change()
        FormMainWindow.minimodewindowopacity = HScrollMiniModeOpacity.Value
        LabelMiniModeOpacityIndicator.Caption = FormMainWindow.minimodewindowopacity & "%"
        Call FormMiniMode.MiniModeAdjustOpacity
    End Sub
    Public Sub HScrollMiniModeOpacity_Scroll()
        Call HScrollMiniModeOpacity_Change
    End Sub
    Public Sub CheckboxMiniModeDarkTheme_Click()
        Select Case FormMainWindow.minimodewindowdarkthemeswitch
            Case True
                FormMainWindow.minimodewindowdarkthemeswitch = False
                CheckboxMiniModeDarkTheme.Value = 0
                FormMiniMode.BackColor = &HF0F0F0
                FormMiniMode.LabelClockHour.ForeColor = &H0
                FormMiniMode.LabelClockDot.ForeColor = &H0
                FormMiniMode.LabelClockMin.ForeColor = &H0
                FormMiniMode.LabelClockSec.ForeColor = &H0
                FormMiniMode.LabelClockDate.ForeColor = &H0
                FormMiniMode.LabelTimerTitle.ForeColor = &H0
                FormMiniMode.LabelTimerDisplay.ForeColor = &H0
                FormMiniMode.LabelLotteryTitle.ForeColor = &H0
                FormMiniMode.LabelLotteryDisplay.ForeColor = &H0
            Case False
                FormMainWindow.minimodewindowdarkthemeswitch = True
                CheckboxMiniModeDarkTheme.Value = 1
                FormMiniMode.BackColor = &H0
                FormMiniMode.LabelClockHour.ForeColor = &HFFFFFF
                FormMiniMode.LabelClockDot.ForeColor = &HFFFFFF
                FormMiniMode.LabelClockMin.ForeColor = &HFFFFFF
                FormMiniMode.LabelClockSec.ForeColor = &HFFFFFF
                FormMiniMode.LabelClockDate.ForeColor = &HFFFFFF
                FormMiniMode.LabelTimerTitle.ForeColor = &HFFFFFF
                FormMiniMode.LabelTimerDisplay.ForeColor = &HFFFFFF
                FormMiniMode.LabelLotteryTitle.ForeColor = &HFFFFFF
                FormMiniMode.LabelLotteryDisplay.ForeColor = &HFFFFFF
        End Select
    End Sub
    Public Sub CheckboxMiniModeAutoExpand_Click()
        Select Case FormMainWindow.minimodeautoexpandswitch
            Case True
                FormMainWindow.minimodeautoexpandswitch = False
                CheckboxMiniModeAutoExpand.Value = 0
            Case False
                FormMainWindow.minimodeautoexpandswitch = True
                CheckboxMiniModeAutoExpand.Value = 1
        End Select
    End Sub
    Public Sub CheckboxMiniModeBlinkWhenMinuteHitsZero_Click()
        Select Case FormMainWindow.minimodeclockoclockblinkswitch
            Case True
                FormMainWindow.minimodeclockoclockblinkswitch = False
                CheckboxMiniModeBlinkWhenMinuteHitsZero.Value = 0
            Case False
                FormMainWindow.minimodeclockoclockblinkswitch = True
                CheckboxMiniModeBlinkWhenMinuteHitsZero.Value = 1
        End Select
    End Sub
    Public Sub CheckboxMiniModeBlinkDots_Click()
        Select Case FormMainWindow.minimodeclockdotblinkswitch
            Case True
                FormMainWindow.minimodeclockdotblinkswitch = False
                CheckboxMiniModeBlinkDots.Value = 0
            Case False
                FormMainWindow.minimodeclockdotblinkswitch = True
                CheckboxMiniModeBlinkDots.Value = 1
        End Select
    End Sub
    Public Sub CheckboxMiniModeShowSeconds_Click()
        Select Case FormMainWindow.minimodeclockshowsecondsswitch
            Case True
                FormMainWindow.minimodeclockshowsecondsswitch = False
                CheckboxMiniModeShowSeconds.Value = 0
                FormMiniMode.LabelClockSec.Visible = False
            Case False
                FormMainWindow.minimodeclockshowsecondsswitch = True
                CheckboxMiniModeShowSeconds.Value = 1
                FormMiniMode.LabelClockSec.Visible = True
        End Select
    End Sub
    Public Sub CheckboxMiniModeAlwaysShowDate_Click()
        Select Case FormMainWindow.minimodeclockalwaysshowdateswitch
            Case True
                FormMainWindow.minimodeclockalwaysshowdateswitch = False
                CheckboxMiniModeAlwaysShowDate.Value = 0
            Case False
                FormMainWindow.minimodeclockalwaysshowdateswitch = True
                CheckboxMiniModeAlwaysShowDate.Value = 1
        End Select
    End Sub
    Public Sub CheckboxMiniModeTimerOverwriteDate_Click()
        Select Case FormMainWindow.minimodetimeroverwritedateswitch
            Case True
                FormMainWindow.minimodetimeroverwritedateswitch = False
                CheckboxMiniModeTimerOverwriteDate.Value = 0
            Case False
                FormMainWindow.minimodetimeroverwritedateswitch = True
                CheckboxMiniModeTimerOverwriteDate.Value = 1
        End Select
        'Update immediately...
        Call FormMainWindow.TimerClock_Timer
    End Sub
    Public Sub CheckboxMiniMode24HrFormat_Click()
        Select Case FormMainWindow.minimodeclock24hrformatswitch
            Case True
                FormMainWindow.minimodeclock24hrformatswitch = False
                CheckboxMiniMode24HrFormat.Value = 0
            Case False
                FormMainWindow.minimodeclock24hrformatswitch = True
                CheckboxMiniMode24HrFormat.Value = 1
        End Select
        'Update immediately...
        Call FormMainWindow.TimerClock_Timer
    End Sub

    'Display settings...
    Public Sub CheckboxDisplayWindowAnimations_Click()
        Select Case FormMainWindow.setanimationswitch
            Case True
                FormMainWindow.setanimationswitch = False
                CheckboxDisplayWindowAnimations.Value = 0
            Case False
                FormMainWindow.setanimationswitch = True
                CheckboxDisplayWindowAnimations.Value = 1
        End Select
    End Sub
    Public Sub CheckboxDisplayUseOldScrollAnimationInLotteryWindow_Click()
        Select Case FormMainWindow.useoldscrollanimationinlotterywindowswitch
            Case True
                FormMainWindow.useoldscrollanimationinlotterywindowswitch = False
                CheckboxDisplayUseOldScrollAnimationInLotteryWindow.Value = 0
            Case False
                FormMainWindow.useoldscrollanimationinlotterywindowswitch = True
                CheckboxDisplayUseOldScrollAnimationInLotteryWindow.Value = 1
        End Select
    End Sub
    Public Sub CheckboxDisplayLightBulbIndicators_Click()
        Select Case FormMainWindow.lightbulbindicatorsswitch
            Case True
                FormMainWindow.lightbulbindicatorsswitch = False
                CheckboxDisplayLightBulbIndicators.Value = 0
                FormMainWindow.ShapeLightTimer.Visible = False
                FormMainWindow.ShapeLightLottery.Visible = False
                FormMainWindow.ShapeLightSoundSwitch.Visible = False
                FormMainWindow.ShapeLightBigFloatingClockSwitch.Visible = False
                FormMainWindow.ShapeLightLotteryGroupSwitch.Visible = False
                FormMiniMode.ShapeLightTimer.Visible = False
                FormMiniMode.ShapeLightLottery.Visible = False
            Case False
                FormMainWindow.lightbulbindicatorsswitch = True
                CheckboxDisplayLightBulbIndicators.Value = 1
                FormMainWindow.ShapeLightTimer.Visible = True
                FormMainWindow.ShapeLightLottery.Visible = True
                FormMainWindow.ShapeLightSoundSwitch.Visible = True
                FormMainWindow.ShapeLightBigFloatingClockSwitch.Visible = True
                FormMainWindow.ShapeLightLotteryGroupSwitch.Visible = True
                FormMiniMode.ShapeLightTimer.Visible = True
                FormMiniMode.ShapeLightLottery.Visible = True
        End Select
    End Sub

    'Sound settings...
    Public Sub CheckboxSoundsTimerTone_Click()
        Select Case FormMainWindow.timertoneswitch
            Case True
                FormMainWindow.timertoneswitch = False
                CheckboxSoundsTimerTone.Value = 0
            Case False
                FormMainWindow.timertoneswitch = True
                CheckboxSoundsTimerTone.Value = 1
        End Select
    End Sub
    Public Sub CheckboxSoundsLotteryTone_Click()
        Select Case FormMainWindow.lotterytoneswitch
            Case True
                FormMainWindow.lotterytoneswitch = False
                CheckboxSoundsLotteryTone.Value = 0
            Case False
                FormMainWindow.lotterytoneswitch = True
                CheckboxSoundsLotteryTone.Value = 1
        End Select
    End Sub
    Public Sub CheckboxSoundsInterfaceSounds_Click()
        Select Case FormMainWindow.interfacesoundswitch
            Case True
                FormMainWindow.interfacesoundswitch = False
                CheckboxSoundsInterfaceSounds.Value = 0
            Case False
                FormMainWindow.interfacesoundswitch = True
                CheckboxSoundsInterfaceSounds.Value = 1
        End Select
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
