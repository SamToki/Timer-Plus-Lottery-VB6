VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer+Lottery　v8.20　by Sam Toki"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   15240
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
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":0CB2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7890
   ScaleWidth      =   15240
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13545
      MouseIcon       =   "FormMainWindow.frx":0E04
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   210
      Width           =   1485
   End
   Begin VB.CommandButton CmdRestartComputer 
      Caption         =   "Restart"
      Height          =   435
      Left            =   11760
      MouseIcon       =   "FormMainWindow.frx":0F56
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Restart Computer"
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton CmdShutDownComputer 
      Caption         =   "Shut Down"
      Height          =   435
      Left            =   10185
      MouseIcon       =   "FormMainWindow.frx":10A8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Shut Down Computer"
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton CmdLockCurrentUser 
      Caption         =   "Lock"
      Height          =   435
      Left            =   8610
      MouseIcon       =   "FormMainWindow.frx":11FA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Lock Current User"
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton CmdRunWindowsCalculator 
      Caption         =   "Calculator"
      Height          =   435
      Left            =   7035
      MouseIcon       =   "FormMainWindow.frx":134C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Run Windows Calculator"
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton CmdBigFloatingClockSwitch 
      Caption         =   "Big Clock: OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4620
      MouseIcon       =   "FormMainWindow.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   210
      Width           =   2220
   End
   Begin VB.CommandButton CmdSoundSwitch 
      Caption         =   "Sound: ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2415
      MouseIcon       =   "FormMainWindow.frx":15F0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   2010
   End
   Begin VB.CommandButton CmdMiniMode 
      Caption         =   "Mini Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      MouseIcon       =   "FormMainWindow.frx":1742
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   210
      Width           =   2010
   End
   Begin VB.Frame FrameTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Timer"
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
      Height          =   6740
      Left            =   210
      TabIndex        =   8
      Top             =   840
      Width           =   6630
      Begin VB.CheckBox CheckboxShutdownWhenTimerEnds 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shut down computer when time is up"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1050
         MaskColor       =   &H000000FF&
         MouseIcon       =   "FormMainWindow.frx":1894
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2415
         Width           =   4320
      End
      Begin VB.Timer TimerTimer 
         Interval        =   500
         Left            =   5670
         Top             =   1365
      End
      Begin VB.CommandButton CmdTimerClear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         MouseIcon       =   "FormMainWindow.frx":19E6
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   5775
         Width           =   2010
      End
      Begin VB.CommandButton CmdTimerReset 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3360
         MouseIcon       =   "FormMainWindow.frx":1B38
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   5040
         Width           =   2010
      End
      Begin VB.CommandButton CmdTimerStartPauseResume 
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1050
         MouseIcon       =   "FormMainWindow.frx":1C8A
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   5040
         Width           =   2010
      End
      Begin VB.CommandButton CmdTimerSecMinus10 
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         MouseIcon       =   "FormMainWindow.frx":1DDC
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   4305
         Width           =   1695
      End
      Begin VB.CommandButton CmdTimerSecPlus1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4515
         MouseIcon       =   "FormMainWindow.frx":1F2E
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   3885
         Width           =   855
      End
      Begin VB.CommandButton CmdTimerMinMinus10 
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1050
         MouseIcon       =   "FormMainWindow.frx":2080
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   4305
         Width           =   1695
      End
      Begin VB.CommandButton CmdTimerMinPlus1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1890
         MouseIcon       =   "FormMainWindow.frx":21D2
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   3885
         Width           =   855
      End
      Begin VB.CommandButton CmdTimerSecMinus1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         MouseIcon       =   "FormMainWindow.frx":2324
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3885
         Width           =   855
      End
      Begin VB.CommandButton CmdTimerMinMinus1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1050
         MouseIcon       =   "FormMainWindow.frx":2476
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   3885
         Width           =   855
      End
      Begin VB.CommandButton CmdTimerSecPlus10 
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         MouseIcon       =   "FormMainWindow.frx":25C8
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   3465
         Width           =   1695
      End
      Begin VB.CommandButton CmdTimerMinPlus10 
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1050
         MouseIcon       =   "FormMainWindow.frx":271A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   3465
         Width           =   1695
      End
      Begin VB.CommandButton CmdTimerSecInput 
         Caption         =   "Enter..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         MouseIcon       =   "FormMainWindow.frx":286C
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   2940
         Width           =   1695
      End
      Begin VB.CommandButton CmdTimerMinInput 
         Caption         =   "Enter..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1050
         MouseIcon       =   "FormMainWindow.frx":29BE
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label LabelTimerDot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   63.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1425
         Left            =   2835
         TabIndex        =   10
         Top             =   340
         Width           =   780
      End
      Begin VB.Label LabelTimerSec 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   63.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1425
         Left            =   3570
         TabIndex        =   11
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label LabelTimerMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   63.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1425
         Left            =   315
         TabIndex        =   9
         Top             =   480
         Width           =   2565
      End
      Begin VB.Shape ShapeLightTimer 
         BackColor       =   &H00666666&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         Height          =   330
         Left            =   5880
         Shape           =   3  'Circle
         Top             =   525
         Width           =   330
      End
      Begin VB.Label LabelTimerEndsAt 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Countdown ends at 00:00:00"
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
         Height          =   375
         Left            =   1050
         TabIndex        =   12
         Top             =   1995
         Width           =   4875
      End
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
      Height          =   6740
      Left            =   7035
      TabIndex        =   27
      Top             =   840
      Width           =   7995
      Begin VB.CommandButton CmdLotteryRepeatTenTimes 
         Caption         =   "REPEAT TEN TIMES"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   315
         MouseIcon       =   "FormMainWindow.frx":2B10
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   5565
         Width           =   7365
      End
      Begin VB.CommandButton CmdLotteryStartLottery 
         Caption         =   "START LOTTERY"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   315
         MouseIcon       =   "FormMainWindow.frx":2C62
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   4515
         Width           =   7365
      End
      Begin VB.CommandButton CmdLotteryGroupSwitch 
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormMainWindow.frx":2DB4
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   2900
         Width           =   1275
      End
      Begin VB.CommandButton CmdLotteryGroupInput 
         Caption         =   "Enter..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5880
         MouseIcon       =   "FormMainWindow.frx":2F06
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   3570
         Width           =   960
      End
      Begin VB.CommandButton CmdLotteryGroupPlus10 
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5040
         MouseIcon       =   "FormMainWindow.frx":3058
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   3570
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryGroupPlus1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         MouseIcon       =   "FormMainWindow.frx":31AA
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   3570
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryGroupMinus1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3780
         MouseIcon       =   "FormMainWindow.frx":32FC
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   3570
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryGroupMinus10 
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormMainWindow.frx":344E
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   3570
         Width           =   645
      End
      Begin VB.TextBox TextboxLotteryGroup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   48
         Text            =   "4"
         Top             =   2940
         Width           =   1800
      End
      Begin VB.CommandButton CmdLotteryTotalInput 
         Caption         =   "Enter..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5880
         MouseIcon       =   "FormMainWindow.frx":35A0
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   1575
         Width           =   960
      End
      Begin VB.CommandButton CmdLotteryTotalPlus10 
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5040
         MouseIcon       =   "FormMainWindow.frx":36F2
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   1575
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryTotalPlus1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         MouseIcon       =   "FormMainWindow.frx":3844
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1575
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryTotalMinus1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3780
         MouseIcon       =   "FormMainWindow.frx":3996
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1575
         Width           =   645
      End
      Begin VB.CommandButton CmdLotteryTotalMinus10 
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormMainWindow.frx":3AE8
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   1575
         Width           =   645
      End
      Begin VB.TextBox TextboxLotteryTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   40
         Text            =   "50"
         Top             =   1000
         Width           =   1800
      End
      Begin VB.Timer TimerLotteryContinuous 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7455
         Top             =   5985
      End
      Begin VB.TextBox TextboxRecord8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   35
         Text            =   "0"
         Top             =   3000
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecordX 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   37
         Text            =   "0"
         Top             =   3720
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   36
         Text            =   "0"
         Top             =   3360
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   34
         Text            =   "0"
         Top             =   2640
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   33
         Text            =   "0"
         Top             =   2280
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   32
         Text            =   "0"
         Top             =   1920
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   31
         Text            =   "0"
         Top             =   1560
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   30
         Text            =   "0"
         Top             =   1200
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         Text            =   "0"
         Top             =   840
         Width           =   2370
      End
      Begin VB.TextBox TextboxRecord1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   2370
      End
      Begin VB.Shape ShapeLightLottery 
         BackColor       =   &H00666666&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   330
         Left            =   7350
         Shape           =   3  'Circle
         Top             =   525
         Width           =   330
      End
      Begin VB.Shape ShapeLightLotteryGroupSwitch 
         BackColor       =   &H00666666&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFCC22&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFAA00&
         Height          =   520
         Left            =   3105
         Top             =   2850
         Width           =   1360
      End
      Begin VB.Label LabelLotteryGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Division:"
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
         Height          =   375
         Left            =   3045
         TabIndex        =   46
         Top             =   2380
         Width           =   2355
      End
      Begin VB.Label LabelLotteryTotal2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "from 1 to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3780
         TabIndex        =   39
         Top             =   1050
         Width           =   1125
      End
      Begin VB.Label LabelLotteryTotal1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Number Range:"
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
         Height          =   375
         Left            =   3045
         TabIndex        =   38
         Top             =   525
         Width           =   2355
      End
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape ShapeLightSoundSwitch 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFCC22&
      FillColor       =   &H00FFAA00&
      FillStyle       =   0  'Solid
      Height          =   520
      Left            =   2370
      Top             =   160
      Width           =   2100
   End
   Begin VB.Shape ShapeLightBigFloatingClockSwitch 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFCC22&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFAA00&
      Height          =   520
      Left            =   4580
      Top             =   160
      Width           =   2310
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   420
      Left            =   2835
      TabIndex        =   56
      Top             =   525
      Visible         =   0   'False
      Width           =   450
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   794
      _cy             =   741
   End
   Begin VB.Menu MenuTimer 
      Caption         =   "&Timer"
      Begin VB.Menu MenuTimerStartPauseResume 
         Caption         =   "Start"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuTimerReset 
         Caption         =   "Reset"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuTimerClear 
         Caption         =   "Clear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuTimer1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuTimerShutdownWhenTimerEnds 
         Caption         =   "Shut down computer when time is up"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuLottery 
      Caption         =   "L&ottery"
      Begin VB.Menu MenuLotteryStartLottery 
         Caption         =   "Start Lottery"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuLotteryRepeatTenTimes 
         Caption         =   "Repeat Ten Times"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenuLottery1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuLotteryGroupSwitch 
         Caption         =   "Group Division"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MenuLottery2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuLotteryClearHistory 
         Caption         =   "Clear History"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu MenuCtrlExtras 
      Caption         =   "Options / &Extras"
      Begin VB.Menu MenuExtrasMiniMode 
         Caption         =   "Mini Mode"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuExtras1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExtrasSoundSwitch 
         Caption         =   "Sound Switch"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuExtras2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExtrasBigFloatingClock 
         Caption         =   "Big Floating Clock"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuExtras3_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExtrasRunWindowsCalculator 
         Caption         =   "Run Windows Calculator"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MenuExtrasLockCurrentUser 
         Caption         =   "Lock Current User"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MenuExtrasShutDownComputer 
         Caption         =   "Shut Down Computer"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MenuExtrasRestartComputer 
         Caption         =   "Restart Computer"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu MenuExtras4_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExtrasHideMainWindow 
         Caption         =   "Hide Main Window (Caution)"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuSettings 
      Caption         =   "&Settings..."
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "&About"
      Begin VB.Menu MenuAboutName 
         Caption         =   "Timer+Lottery"
      End
      Begin VB.Menu MenuAboutVersion 
         Caption         =   "v8.20 Release Version　|　for Windows 7,8,10　|　English (US)"
      End
      Begin VB.Menu MenuAboutDate 
         Caption         =   "Last compiled on Fri, Sep 25, 2020"
      End
      Begin VB.Menu MenuAboutFirst 
         Caption         =   "First version built on Fri, Mar 24, 2017"
      End
      Begin VB.Menu MenuAbout1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutAuthor 
         Caption         =   "Author: Sam Toki"
      End
      Begin VB.Menu MenuAboutOrganization 
         Caption         =   "Organization: SAM TOKI STUDIO"
      End
      Begin VB.Menu MenuAboutFrom 
         Caption         =   "From: Xidian University, China"
      End
      Begin VB.Menu MenuAboutContact 
         Caption         =   "Contact: SamToki@outlook.com"
      End
      Begin VB.Menu MenuAbout2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCopyright 
         Caption         =   "TM ＆ (C) 2015-2020 SAM TOKI STUDIO. All rights reserved."
      End
      Begin VB.Menu MenuAboutTrademark 
         Caption         =   "SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries."
      End
      Begin VB.Menu MenuAbout3_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCommercial 
         Caption         =   "Commercial use of this software is strictly prohibited."
      End
   End
   Begin VB.Menu Menu2_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Ａ字あ (&L)"
      Begin VB.Menu MenuLanguageENG 
         Caption         =   "English (United States)"
         Checked         =   -1  'True
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MenuLanguageCHS 
         Caption         =   "中文（简体）"
         Enabled         =   0   'False
         Shortcut        =   +{F2}
      End
      Begin VB.Menu MenuLanguageCHT 
         Caption         =   "中文（繁體）"
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu MenuLanguageJPN 
         Caption         =   "日本語"
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu Menu3_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuEXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  Timer+Lottery
'
'  Powered by Sam Toki
'  Version: v8.20 Release Version ENG
'  Date:    09/25/2020 (Fri.)
'  History: First version v0.10 Beta was built on 03/24/2017.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2020 SAM TOKI STUDIO. All rights reserved.
'             SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public setlanguage As String

'Declare Controls...
Public timerswitch As Boolean
Public shutdownwhentimeupswitch As Boolean
Public lotterygroupswitch As Boolean
Public bigfloatingclockswitch As Boolean

'Declare SettingsFeatures...
Public lotterypreventrepeatswitch As Boolean
Public lotterywindowdarkthemeswitch As Boolean
Public minimodewindowopacity As Integer
Public minimodewindowdarkthemeswitch As Boolean
Public minimodeautohidesettimeout As Integer
Public minimodeautohidetimeout As Integer
Public minimodeautoexpandswitch As Boolean
Public minimodeclockoclockblinkswitch As Boolean
Public minimodeclockdotblinkswitch As Boolean
Public minimodeclockshowsecondsswitch As Boolean
Public minimodeclockalwaysshowdateswitch As Boolean
Public minimodetimeroverwritedateswitch As Boolean
Public minimodeclock24hrformatswitch As Boolean
Public bigfloatingclockshadowswitch As Boolean
Public bigfloatingclockautohideswitch As Boolean
Public bigfloatingclockoclockblinkswitch As Boolean
Public bigfloatingclock24hrformatswitch As Boolean
Public shutdowncountdowntype As String
Public shutdowncountdowntimeout As Integer

'Declare SettingsDisplay...
Public setanimationswitch As Boolean
Public useoldscrollanimationinlotterywindowswitch As Boolean
Public lightbulbindicatorsswitch As Boolean

'Declare SettingsSounds...
Public soundswitch As Boolean
Public timertoneswitch As Boolean
Public lotterytoneswitch As Boolean
Public interfacesoundswitch As Boolean

'Declare Timer...
Public timersettime As Long
Public timersettimemin As Long
Public timersettimesec As Long
Public timercountdowntime As Long
Public timercountdowntimemin As Long
Public timercountdowntimesec As Long
Public timerendtime As Long
Public timerendtimehour As Long
Public timerendtimemin As Long
Public timerendtimesec As Long
Public timerendtimehourtext As String
Public timerendtimemintext As String
Public timerendtimesectext As String
Public timerexpiredsec As Long

'Declare Lottery...
Public lotterytotal As Integer
Public lotterygroup As Integer
Public lotterynumber As Integer
Public lotterynumberrecordX As Integer
Public lotterynumberrecord9 As Integer
Public lotterynumberrecord8 As Integer
Public lotterynumberrecord7 As Integer
Public lotterynumberrecord6 As Integer
Public lotterynumberrecord5 As Integer
Public lotterynumberrecord4 As Integer
Public lotterynumberrecord3 As Integer
Public lotterynumberrecord2 As Integer
Public lotterynumberrecord1 As Integer
Public lotteryhinttext As String
Public lotteryscrolltext As String
Public lotteryscrollprogress As Integer
Public lotterytimeout As Integer
Public lotterylooper As Integer

'Declare Clock...
Public clockhour As Long
Public clockmin As Long
Public clocksec As Long
Public clockmonth As Integer
Public clockday As Integer
Public clockweekday As String

'Declare Dialog...
Public inputnumbermode As String
Public inputnumberdigits As Integer
Public answer

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Public Sub Form_Load()
        'FIRST GENERAL RESET

        setlanguage = "ENG"

        timerswitch = False
        shutdownwhentimeupswitch = False
        lotterygroupswitch = False
        bigfloatingclockswitch = False

        lotterypreventrepeatswitch = True
        lotterywindowdarkthemeswitch = False
        minimodewindowopacity = 80
        minimodewindowdarkthemeswitch = True
        minimodeautohidesettimeout = 5
        minimodeautohidetimeout = 5
        minimodeautoexpandswitch = True
        minimodeclockoclockblinkswitch = True
        minimodeclockdotblinkswitch = False
        minimodeclockshowsecondsswitch = True
        minimodeclockalwaysshowdateswitch = True
        minimodetimeroverwritedateswitch = True
        minimodeclock24hrformatswitch = True
        bigfloatingclockshadowswitch = True
        bigfloatingclockautohideswitch = True
        bigfloatingclockoclockblinkswitch = True
        bigfloatingclock24hrformatswitch = True
        shutdowncountdowntype = "Shutdown"
        shutdowncountdowntimeout = 16

        setanimationswitch = True
        useoldscrollanimationinlotterywindowswitch = False
        lightbulbindicatorsswitch = True

        soundswitch = True
        timertoneswitch = True
        lotterytoneswitch = True
        interfacesoundswitch = True
        
        timersettime = 0
        timersettimemin = 0
        timersettimesec = 0
        timercountdowntime = 0
        timercountdowntimemin = 0
        timercountdowntimesec = 0
        timerendtime = 0
        timerendtimehour = 0
        timerendtimemin = 0
        timerendtimesec = 0
        timerendtimehourtext = "00"
        timerendtimemintext = "00"
        timerendtimesectext = "00"
        timerexpiredsec = 0
        
        lotterytotal = 50
        lotterygroup = 4
        lotterynumber = 0
        lotterynumberrecordX = -1
        lotterynumberrecord9 = -1
        lotterynumberrecord8 = -1
        lotterynumberrecord7 = -1
        lotterynumberrecord6 = -1
        lotterynumberrecord5 = -1
        lotterynumberrecord4 = -1
        lotterynumberrecord3 = -1
        lotterynumberrecord2 = -1
        lotterynumberrecord1 = -1
        lotteryhinttext = "HintText Abg"
        lotteryscrolltext = "ScrText Abg"
        lotteryscrollprogress = 0
        lotterytimeout = -1
        lotterylooper = 10

        clockhour = 0
        clockmin = 0
        clocksec = 0
        clockmonth = 0
        clockday = 0
        clockweekday = "??"
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    Public Sub TimerClock_Timer()
        clockhour = Hour(Time)
        clockmin = Minute(Time)
        clocksec = Second(Time)
        clockmonth = Month(Date)
        clockday = Day(Date)
        Select Case Weekday(Date)
            Case 1
                clockweekday = "Sun"
            Case 2
                clockweekday = "Mon"
            Case 3
                clockweekday = "Tue"
            Case 4
                clockweekday = "Wed"
            Case 5
                clockweekday = "Thu"
            Case 6
                clockweekday = "Fri"
            Case 7
                clockweekday = "Sat"
        End Select

        Select Case minimodeclock24hrformatswitch
            Case True
                FormMiniMode.LabelClockHour.Caption = Format(Hour(Time), "00")
            Case False
                Select Case Hour(Time)
                    Case 0
                        FormMiniMode.LabelClockHour.Caption = "12"
                    Case Is >= 1
                        If Hour(Time) > 12 Then FormMiniMode.LabelClockHour.Caption = Hour(Time) - 12 Else FormMiniMode.LabelClockHour.Caption = Hour(Time)
                End Select
        End Select
        FormMiniMode.LabelClockMin.Caption = Format(Minute(Time), "00")
        FormMiniMode.LabelClockSec.Caption = Format(Second(Time), "00")
        If (minimodetimeroverwritedateswitch = True And timerswitch = True) Then
            FormMiniMode.LabelClockDate.Caption = LabelTimerMin.Caption & " : " & LabelTimerSec.Caption
        Else
            FormMiniMode.LabelClockDate.Caption = clockmonth & "/" & clockday & " (" & clockweekday & ")"
        End If

        If minimodeclockoclockblinkswitch = True And Minute(Time) = 0 And Second(Time) = 0 And FormMiniMode.minimodeclockoclockblinkrepeatedtimes = -1 Then
            'CAUTION: TEST ONLY!!
            'If minimodeclockoclockblinkswitch = True And Second(Time) = 0 And FormMiniMode.minimodeclockoclockblinkrepeatedtimes = -1 Then
            FormMiniMode.minimodeclockoclockblinkrepeatedtimes = 0
            FormMiniMode.minimodeclockoclockblinkopacity = 255 * (minimodewindowopacity / 100)
            FormMiniMode.minimodeclockoclockblinksetopacity = FormMiniMode.minimodeclockoclockblinkopacity
            FormMiniMode.TimerMiniModeOclockBlink.Enabled = True
        End If
    End Sub

    Public Sub TimerTimer_Timer()
        If (timerswitch = True And ShapeLightTimer.FillStyle = 1) Then
            ShapeLightTimer.BorderStyle = 1
            ShapeLightTimer.FillStyle = 0
            FormMiniMode.ShapeLightTimer.BorderStyle = 1
            FormMiniMode.ShapeLightTimer.FillStyle = 0
        Else
            ShapeLightTimer.BorderStyle = 0
            ShapeLightTimer.FillStyle = 1
            FormMiniMode.ShapeLightTimer.BorderStyle = 0
            FormMiniMode.ShapeLightTimer.FillStyle = 1
        End If
        If timersettime <= 0 Then
            MenuTimerStartPauseResume.Enabled = False
            CmdTimerStartPauseResume.Enabled = False
            FormMiniMode.CmdTimerStartPauseResume.Enabled = False
        Else
            MenuTimerStartPauseResume.Enabled = True
            CmdTimerStartPauseResume.Enabled = True
            FormMiniMode.CmdTimerStartPauseResume.Enabled = True
        End If

        'Prevent adjusting time when timer is ongoing.
        If timersettime = timercountdowntime Then
            CmdTimerMinInput.Enabled = True
            CmdTimerMinMinus1.Enabled = True
            CmdTimerMinMinus10.Enabled = True
            CmdTimerMinPlus1.Enabled = True
            CmdTimerMinPlus10.Enabled = True
            CmdTimerSecInput.Enabled = True
            CmdTimerSecMinus1.Enabled = True
            CmdTimerSecMinus10.Enabled = True
            CmdTimerSecPlus1.Enabled = True
            CmdTimerSecPlus10.Enabled = True
        Else
            CmdTimerMinInput.Enabled = False
            CmdTimerMinMinus1.Enabled = False
            CmdTimerMinMinus10.Enabled = False
            CmdTimerMinPlus1.Enabled = False
            CmdTimerMinPlus10.Enabled = False
            CmdTimerSecInput.Enabled = False
            CmdTimerSecMinus1.Enabled = False
            CmdTimerSecMinus10.Enabled = False
            CmdTimerSecPlus1.Enabled = False
            CmdTimerSecPlus10.Enabled = False
        End If

        'Timer Display
        Select Case timerswitch
            Case True
                timercountdowntime = timerendtime - (clockhour * 3600 + clockmin * 60 + clocksec)
            Case False
                timerendtime = (clockhour * 3600 + clockmin * 60 + clocksec) + timercountdowntime
        End Select

        timercountdowntimemin = Int(timercountdowntime \ 60)
        timercountdowntimesec = timercountdowntime Mod 60
        If timercountdowntimemin < 0 Then timercountdowntimemin = 0
        If timercountdowntimesec < 0 Then timercountdowntimesec = 0
        timerendtimehour = Int(timerendtime \ 3600)
        timerendtimemin = Int((timerendtime Mod 3600) \ 60)
        timerendtimesec = timerendtime Mod 60
        LabelTimerMin.Caption = timercountdowntimemin
        LabelTimerSec.Caption = Format(timercountdowntimesec, "00")
        FormMiniMode.LabelTimerDisplay.Caption = LabelTimerMin.Caption & ":" & LabelTimerSec.Caption
        timerendtimehourtext = Format(timerendtimehour Mod 24, "00")
        timerendtimemintext = Format(timerendtimemin, "00")
        timerendtimesectext = Format(timerendtimesec, "00")

        LabelTimerEndsAt.Caption = "Countdown ends at " & timerendtimehourtext & ":" & timerendtimemintext & ":" & timerendtimesectext
        Select Case timerswitch
            Case True
                CmdTimerStartPauseResume.Caption = "PAUSE"
                MenuTimerStartPauseResume.Caption = "Pause"
                FormMiniMode.CmdTimerStartPauseResume.Caption = "PAUSE"
            Case False
                If timersettime = timercountdowntime Then
                    CmdTimerStartPauseResume.Caption = "START"
                    MenuTimerStartPauseResume.Caption = "Start"
                    FormMiniMode.CmdTimerStartPauseResume.Caption = "START"
                Else
                    CmdTimerStartPauseResume.Caption = "RESUME"
                    MenuTimerStartPauseResume.Caption = "Resume"
                    FormMiniMode.CmdTimerStartPauseResume.Caption = "RESUME"
                End If
        End Select

        'Interface sound...
        If (soundswitch = True And (Not (FormMainWindow.WindowState = 1)) And timerswitch = True And TimerLotteryContinuous.Enabled = False) Then
            Select Case interfacesoundswitch
                Case True
                    WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Menu Command.wav"
                Case False
                    WindowsMediaPlayer1.URL = ""
            End Select
        End If

        'Time up...
        If timercountdowntime < 0 Then
            If FormTimeUp.TimerExpiredTimeCount.Enabled = True Then Exit Sub
            timerswitch = False
            timerexpiredsec = 0
            Call FormTimeUp.TimerExpiredTimeCount_Timer
            FormTimeUp.TimerExpiredTimeCount.Enabled = True
            FormTimeUp.TimerTimeUpTextBlink.Enabled = True
            FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
            FormTimeUp.Move ((Screen.Width / 2) - (12510 / 2)), -2745, 12510, 0
            FormTimeUp.windowanimationtargetleft = (Screen.Width / 2) - (12510 / 2)
            FormTimeUp.windowanimationtargettop = (Screen.Height / 2) - (2745 / 2)
            FormTimeUp.windowanimationtargetwidth = 12510
            FormTimeUp.windowanimationtargetheight = 2745
            FormTimeUp.Show
            'Shut down computer when time is up...
            If shutdownwhentimeupswitch = True Then
                shutdowncountdowntype = "Shutdown"
                shutdowncountdowntimeout = 16
                FormShutdownCountdown.LabelHinttextA.Caption = "Your computer is set to shut down when time is up."
                FormShutdownCountdown.LabelHinttextB.Caption = "Please save your files in time."
                FormShutdownCountdown.CmdOK.Caption = "Shut Down Now"
                FormShutdownCountdown.TimerShutdownCountdown.Enabled = True
                FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
                FormShutdownCountdown.Move ((Screen.Width / 2) - (12510 / 2)), -2745, 12510, 0
                FormShutdownCountdown.windowanimationtargetleft = (Screen.Width / 2) - (12510 / 2)
                FormShutdownCountdown.windowanimationtargettop = 0
                FormShutdownCountdown.windowanimationtargetwidth = 12510
                FormShutdownCountdown.windowanimationtargetheight = 2745
                FormShutdownCountdown.Show
            End If
        End If
    End Sub

    Public Sub TimerLotteryContinuous_Timer()
        If lotterylooper <= 0 Then TimerLotteryContinuous.Enabled = False:  lotterylooper = 10: FormMainWindow.ShapeLightLottery.BorderStyle = 0: FormMainWindow.ShapeLightLottery.FillStyle = 1: FormMiniMode.ShapeLightLottery.BorderStyle = 0: FormMiniMode.ShapeLightLottery.FillStyle = 1: FormMainWindow.MousePointer = 99: CmdLotteryStartLottery.Enabled = True: CmdLotteryRepeatTenTimes.Enabled = True: Exit Sub
        Call LotteryExecuteOnce
        If lotterylooper > 0 Then lotterylooper = lotterylooper - 1
    End Sub

'[] LOTTERY []

    Public Sub LotteryExecuteOnce()
LABEL_LotteryExecuteOnce_RANDOM_NUMBER_REGENERATE:
        Call RandomNumberGenerator
        Select Case lotterypreventrepeatswitch
            Case True
                If (lotterynumber <> lotterynumberrecord1) And (lotterynumber <> lotterynumberrecord2) And (lotterynumber <> lotterynumberrecord3) And (lotterynumber <> lotterynumberrecord4) And (lotterynumber <> lotterynumberrecord5) And (lotterynumber <> lotterynumberrecord6) And (lotterynumber <> lotterynumberrecord7) And (lotterynumber <> lotterynumberrecord8) And (lotterynumber <> lotterynumberrecord9) And (lotterynumber <> lotterynumberrecordX) Then
                    lotterynumberrecordX = lotterynumberrecord9
                    lotterynumberrecord9 = lotterynumberrecord8
                    lotterynumberrecord8 = lotterynumberrecord7
                    lotterynumberrecord7 = lotterynumberrecord6
                    lotterynumberrecord6 = lotterynumberrecord5
                    lotterynumberrecord5 = lotterynumberrecord4
                    lotterynumberrecord4 = lotterynumberrecord3
                    lotterynumberrecord3 = lotterynumberrecord2
                    lotterynumberrecord2 = lotterynumberrecord1
                    lotterynumberrecord1 = lotterynumber
                    Call LotteryRecordsRefresher
                Else
                    'BUG FIX!! PREVENT THE INFINITE LOOP
                    If lotterytotal > 10 Then
                        GoTo LABEL_LotteryExecuteOnce_RANDOM_NUMBER_REGENERATE
                    Else
                        lotterynumberrecordX = lotterynumberrecord9
                        lotterynumberrecord9 = lotterynumberrecord8
                        lotterynumberrecord8 = lotterynumberrecord7
                        lotterynumberrecord7 = lotterynumberrecord6
                        lotterynumberrecord6 = lotterynumberrecord5
                        lotterynumberrecord5 = lotterynumberrecord4
                        lotterynumberrecord4 = lotterynumberrecord3
                        lotterynumberrecord3 = lotterynumberrecord2
                        lotterynumberrecord2 = lotterynumberrecord1
                        lotterynumberrecord1 = lotterynumber
                        Call LotteryRecordsRefresher
                    End If
                End If
            Case False
                    lotterynumberrecordX = lotterynumberrecord9
                    lotterynumberrecord9 = lotterynumberrecord8
                    lotterynumberrecord8 = lotterynumberrecord7
                    lotterynumberrecord7 = lotterynumberrecord6
                    lotterynumberrecord6 = lotterynumberrecord5
                    lotterynumberrecord5 = lotterynumberrecord4
                    lotterynumberrecord4 = lotterynumberrecord3
                    lotterynumberrecord3 = lotterynumberrecord2
                    lotterynumberrecord2 = lotterynumberrecord1
                    lotterynumberrecord1 = lotterynumber
                    Call LotteryRecordsRefresher
        End Select
    End Sub

    Public Sub RandomNumberGenerator()
        Randomize
        lotterynumber = Int(lotterytotal * Rnd)
    End Sub

    Public Sub LotteryRecordsRefresher()
        Select Case lotterygroupswitch
            Case False
                TextboxRecordX.Text = lotterynumberrecordX + 1
                TextboxRecord9.Text = lotterynumberrecord9 + 1
                TextboxRecord8.Text = lotterynumberrecord8 + 1
                TextboxRecord7.Text = lotterynumberrecord7 + 1
                TextboxRecord6.Text = lotterynumberrecord6 + 1
                TextboxRecord5.Text = lotterynumberrecord5 + 1
                TextboxRecord4.Text = lotterynumberrecord4 + 1
                TextboxRecord3.Text = lotterynumberrecord3 + 1
                TextboxRecord2.Text = lotterynumberrecord2 + 1
                TextboxRecord1.Text = lotterynumberrecord1 + 1
                FormMiniMode.LabelLotteryDisplay.Caption = lotterynumberrecord1 + 1
            Case True
                TextboxRecordX.Text = (Int(lotterynumberrecordX / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecordX Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord9.Text = (Int(lotterynumberrecord9 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord9 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord8.Text = (Int(lotterynumberrecord8 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord8 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord7.Text = (Int(lotterynumberrecord7 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord7 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord6.Text = (Int(lotterynumberrecord6 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord6 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord5.Text = (Int(lotterynumberrecord5 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord5 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord4.Text = (Int(lotterynumberrecord4 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord4 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord3.Text = (Int(lotterynumberrecord3 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord3 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord2.Text = (Int(lotterynumberrecord2 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord2 Mod Int(lotterytotal / lotterygroup) + 1)
                TextboxRecord1.Text = (Int(lotterynumberrecord1 / Int(lotterytotal / lotterygroup)) + 1) & "-" & (lotterynumberrecord1 Mod Int(lotterytotal / lotterygroup) + 1)
                FormMiniMode.LabelLotteryDisplay.Caption = TextboxRecord1.Text
        End Select
    End Sub

    Public Sub LotterySettingsRefresher()
        TextboxLotteryTotal.Text = lotterytotal
        TextboxLotteryGroup.Text = lotterygroup
    End Sub

'[] COMMANDS []

    'CMD Language...
    Public Sub MenuLanguageENG_Click()
        'Call ModuleLoadLanguage.LoadLanguageENG
    End Sub
    Public Sub MenuLanguageCHS_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHS
    End Sub
    Public Sub MenuLanguageCHT_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHT
    End Sub
    Public Sub MenuLanguageJPN_Click()
        'Call ModuleLoadLanguage.LoadLanguageJPN
    End Sub

    'CMD Controls...
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub CmdEXIT_Click()
        Call MenuEXIT_Click
    End Sub
    Public Sub MenuSettings_Click()
        FormSettings.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormSettings.windowanimationtargetleft = (Screen.Width / 2) - (12930 / 2)
        FormSettings.windowanimationtargettop = (Screen.Height / 2) - (6945 / 2)
        FormSettings.windowanimationtargetwidth = 12930
        FormSettings.windowanimationtargetheight = 6945
        FormSettings.Show
    End Sub

    Public Sub MenuTimerStartPauseResume_Click()
        Select Case timerswitch
            Case True
                timerswitch = False
            Case False
                timerswitch = True
        End Select

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdTimerStartPauseResume_Click()
        Call MenuTimerStartPauseResume_Click
    End Sub
    Public Sub MenuTimerReset_Click()
        timerswitch = False
        timercountdowntime = timersettime
        Call TimerTimer_Timer

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Startup.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdTimerReset_Click()
        Call MenuTimerReset_Click
    End Sub
    Public Sub MenuTimerClear_Click()
        timerswitch = False
        timersettime = 0
        timercountdowntime = timersettime
        Call TimerTimer_Timer

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Recycle.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdTimerClear_Click()
        Call MenuTimerClear_Click
    End Sub
    Public Sub MenuTimerShutdownWhenTimerEnds_Click()
        Select Case shutdownwhentimeupswitch
            Case True
                shutdownwhentimeupswitch = False
                CheckboxShutdownWhenTimerEnds.Value = 0
                MenuTimerShutdownWhenTimerEnds.Checked = False
            Case False
                shutdownwhentimeupswitch = True
                CheckboxShutdownWhenTimerEnds.Value = 1
                MenuTimerShutdownWhenTimerEnds.Checked = True
        End Select
    End Sub
    Public Sub CheckboxShutdownWhenTimerEnds_Click()
        Call MenuTimerShutdownWhenTimerEnds_Click
    End Sub
    Public Sub MenuLotteryStartLottery_Click()
        lotteryhinttext = "Now Loading..."
        lotteryscrolltext = "N/A"

        lotteryscrollprogress = 0
        FormLottery.LabelHintText = lotteryhinttext
        FormLottery.LabelScrollText = lotteryscrolltext
        FormLottery.TimerScroll.Enabled = True
        FormLottery.TimerLottery.Enabled = True
        FormLottery.CmdCancel.Enabled = True
        FormMainWindow.ShapeLightLottery.BorderStyle = 1
        FormMainWindow.ShapeLightLottery.FillStyle = 0
        FormMiniMode.ShapeLightLottery.BorderStyle = 1
        FormMiniMode.ShapeLightLottery.FillStyle = 0
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False: FormLottery.Show

        If soundswitch = False Then Exit Sub
        Select Case lotterytoneswitch
            Case True
                WindowsMediaPlayer1.URL = App.Path & "\CZJST_appdata\CZJST_audio\CZJSTaudio_HaiyoreNyarukosanA.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdLotteryStartLottery_Click()
        Call MenuLotteryStartLottery_Click
    End Sub
    Public Sub MenuLotteryRepeatTenTimes_Click()
        CmdLotteryRepeatTenTimes.Enabled = False
        FormMainWindow.MousePointer = 11
        lotterylooper = 10
        TimerLotteryContinuous.Enabled = True
        FormMainWindow.ShapeLightLottery.BorderStyle = 1
        FormMainWindow.ShapeLightLottery.FillStyle = 0
        FormMiniMode.ShapeLightLottery.BorderStyle = 1
        FormMiniMode.ShapeLightLottery.FillStyle = 0

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Connection.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdLotteryRepeatTenTimes_Click()
        Call MenuLotteryRepeatTenTimes_Click
    End Sub
    Public Sub MenuLotteryClearHistory_Click()
        lotterynumberrecordX = -1
        lotterynumberrecord9 = -1
        lotterynumberrecord8 = -1
        lotterynumberrecord7 = -1
        lotterynumberrecord6 = -1
        lotterynumberrecord5 = -1
        lotterynumberrecord4 = -1
        lotterynumberrecord3 = -1
        lotterynumberrecord2 = -1
        lotterynumberrecord1 = -1
        Call LotteryRecordsRefresher

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Recycle.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub MenuLotteryGroupSwitch_Click()
        Select Case lotterygroupswitch
            Case True
                lotterygroupswitch = False
                Call LotteryRecordsRefresher
                MenuLotteryGroupSwitch.Checked = False
                ShapeLightLotteryGroupSwitch.BorderStyle = 0
                ShapeLightLotteryGroupSwitch.FillStyle = 1
                CmdLotteryGroupSwitch.Caption = "OFF"
            Case False
                lotterygroupswitch = True
                Call LotteryRecordsRefresher
                MenuLotteryGroupSwitch.Checked = True
                ShapeLightLotteryGroupSwitch.BorderStyle = 1
                ShapeLightLotteryGroupSwitch.FillStyle = 0
                CmdLotteryGroupSwitch.Caption = "ON"
        End Select
    End Sub
    Public Sub CmdLotteryGroupSwitch_Click()
        Call MenuLotteryGroupSwitch_Click
    End Sub

    Public Sub MenuExtrasMiniMode_Click()
        FormMiniMode.Move 0, 0, 0, 0
        FormMiniMode.windowanimationtargetleft = 0
        FormMiniMode.windowanimationtargettop = 0
        FormMiniMode.windowanimationtargetwidth = 4455
        FormMiniMode.windowanimationtargetheight = 1830
        FormMiniMode.Show

        FormMiniMode.TimerMiniModeAutoHide.Enabled = True
        minimodeautohidetimeout = 10

        If bigfloatingclockswitch = True Then Call MenuExtrasBigFloatingClock_Click

        FormMainWindow.WindowState = 1
    End Sub
    Public Sub CmdMiniMode_Click()
        Call MenuExtrasMiniMode_Click
    End Sub
    Public Sub MenuExtrasSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                MenuExtrasSoundSwitch.Checked = False
                ShapeLightSoundSwitch.BorderStyle = 0
                ShapeLightSoundSwitch.FillStyle = 1
                CmdSoundSwitch.Caption = "Sound: OFF"
            Case False
                soundswitch = True
                MenuExtrasSoundSwitch.Checked = True
                ShapeLightSoundSwitch.BorderStyle = 1
                ShapeLightSoundSwitch.FillStyle = 0
                CmdSoundSwitch.Caption = "Sound: ON"
        End Select
    End Sub
    Public Sub CmdSoundSwitch_Click()
        Call MenuExtrasSoundSwitch_Click
    End Sub
    Public Sub MenuExtrasBigFloatingClock_Click()
        Select Case bigfloatingclockswitch
            Case True
                bigfloatingclockswitch = False
                FormBigFloatingClock.windowanimationtargetwidth = 0
                FormBigFloatingClock.windowanimationtargetheight = 0
                FormBigFloatingClock.bigfloatingclockautohidetimeout = -1
                MenuExtrasBigFloatingClock.Checked = False
                ShapeLightBigFloatingClockSwitch.BorderStyle = 0
                ShapeLightBigFloatingClockSwitch.FillStyle = 1
                CmdBigFloatingClockSwitch.Caption = "Big Clock: OFF"
            Case False
                bigfloatingclockswitch = True
                FormBigFloatingClock.Move (FormSettings.HScrollBigFloatingClockPositionX.Value / 1000) * (Screen.Width - 3060), (FormSettings.VScrollBigFloatingClockPositionY.Value / 1000) * (Screen.Height - 1485), 0, 0
                FormBigFloatingClock.windowanimationtargetwidth = 3060
                FormBigFloatingClock.windowanimationtargetheight = 1485
                FormBigFloatingClock.Show
                FormBigFloatingClock.bigfloatingclockautohidetimeout = -1
                MenuExtrasBigFloatingClock.Checked = True
                ShapeLightBigFloatingClockSwitch.BorderStyle = 1
                ShapeLightBigFloatingClockSwitch.FillStyle = 0
                CmdBigFloatingClockSwitch.Caption = "Big Clock: ON"
        End Select
    End Sub
    Public Sub CmdBigFloatingClockSwitch_Click()
        Call MenuExtrasBigFloatingClock_Click
    End Sub
    Public Sub MenuExtrasRunWindowsCalculator_Click()
        Shell "cmd.exe /c calc", vbHide
    End Sub
    Public Sub CmdRunWindowsCalculator_Click()
        Call MenuExtrasRunWindowsCalculator_Click
    End Sub
    Public Sub MenuExtrasLockCurrentUser_Click()
        Shell "cmd.exe /c rundll32 user32.dll, LockWorkStation", vbHide

        'Interface sound...
        If soundswitch = False Then Exit Sub
        Select Case interfacesoundswitch
            Case True
                WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Logoff Sound.wav"
            Case False
                WindowsMediaPlayer1.URL = ""
        End Select
    End Sub
    Public Sub CmdLockCurrentUser_Click()
        Call MenuExtrasLockCurrentUser_Click
    End Sub
    Public Sub MenuExtrasShutDownComputer_Click()
        shutdowncountdowntype = "Shutdown"
        shutdowncountdowntimeout = 16
        FormShutdownCountdown.LabelHinttextA.Caption = "Your computer is about to shut down."
        FormShutdownCountdown.LabelHinttextB.Caption = "Please save your files in time."
        FormShutdownCountdown.CmdOK.Caption = "Shut Down Now"
        FormShutdownCountdown.TimerShutdownCountdown.Enabled = True
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False

        FormShutdownCountdown.Move ((Screen.Width / 2) - (12510 / 2)), -2745, 12510, 0
        FormShutdownCountdown.windowanimationtargetleft = (Screen.Width / 2) - (12510 / 2)
        FormShutdownCountdown.windowanimationtargettop = 0
        FormShutdownCountdown.windowanimationtargetwidth = 12510
        FormShutdownCountdown.windowanimationtargetheight = 2745
        FormShutdownCountdown.Show
    End Sub
    Public Sub CmdShutDownComputer_Click()
        Call MenuExtrasShutDownComputer_Click
    End Sub
    Public Sub MenuExtrasRestartComputer_Click()
        shutdowncountdowntype = "Restart"
        shutdowncountdowntimeout = 16
        FormShutdownCountdown.LabelHinttextA.Caption = "Your computer is about to restart."
        FormShutdownCountdown.LabelHinttextB.Caption = "Please save your files in time."
        FormShutdownCountdown.CmdOK.Caption = "Restart Now"
        FormShutdownCountdown.TimerShutdownCountdown.Enabled = True
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False

        FormShutdownCountdown.Move ((Screen.Width / 2) - (12510 / 2)), -2745, 12510, 0
        FormShutdownCountdown.windowanimationtargetleft = (Screen.Width / 2) - (12510 / 2)
        FormShutdownCountdown.windowanimationtargettop = 0
        FormShutdownCountdown.windowanimationtargetwidth = 12510
        FormShutdownCountdown.windowanimationtargetheight = 2745
        FormShutdownCountdown.Show
    End Sub
    Public Sub CmdRestartComputer_Click()
        Call MenuExtrasRestartComputer_Click
    End Sub
    Public Sub MenuExtrasHideMainWindow_Click()
        FormMainWindow.Hide
    End Sub

    'CMD Adjustments...
    Public Sub CmdTimerMinInput_Click()
        inputnumbermode = "TimerMin"
        inputnumberdigits = 3
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">"
        FormInputNumber.LabelInputNumber2.Caption = ">"
        FormInputNumber.LabelInputNumber3.Caption = ">"
        FormInputNumber.LabelInputNumber4.Caption = ""
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
        FormInputNumber.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
        FormInputNumber.Show
    End Sub
    Public Sub CmdTimerSecInput_Click()
        inputnumbermode = "TimerSec"
        inputnumberdigits = 2
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">"
        FormInputNumber.LabelInputNumber2.Caption = ">"
        FormInputNumber.LabelInputNumber3.Caption = ""
        FormInputNumber.LabelInputNumber4.Caption = ""
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
        FormInputNumber.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
        FormInputNumber.Show
    End Sub
    Public Sub CmdLotteryTotalInput_Click()
        inputnumbermode = "LotteryTotal"
        inputnumberdigits = 4
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">"
        FormInputNumber.LabelInputNumber2.Caption = ">"
        FormInputNumber.LabelInputNumber3.Caption = ">"
        FormInputNumber.LabelInputNumber4.Caption = ">"
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
        FormInputNumber.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
        FormInputNumber.Show
    End Sub
    Public Sub CmdLotteryGroupInput_Click()
        inputnumbermode = "LotteryGroup"
        inputnumberdigits = 2
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">"
        FormInputNumber.LabelInputNumber2.Caption = ">"
        FormInputNumber.LabelInputNumber3.Caption = ""
        FormInputNumber.LabelInputNumber4.Caption = ""
        FormMainWindow.Enabled = False: FormMiniMode.Enabled = False
        FormInputNumber.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
        FormInputNumber.Show
    End Sub

    Public Sub CmdTimerMinPlus10_Click()
        timersettime = timersettime + 600
        If timersettime >= 60000 Then timersettime = 59999
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerMinPlus1_Click()
        timersettime = timersettime + 60
        If timersettime >= 60000 Then timersettime = 59999
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerMinMinus1_Click()
        timersettime = timersettime - 60
        If timersettime < 0 Then timersettime = 0
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerMinMinus10_Click()
        timersettime = timersettime - 600
        If timersettime < 0 Then timersettime = 0
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerSecPlus10_Click()
        timersettime = timersettime + 10
        If timersettime >= 60000 Then timersettime = 59999
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerSecPlus1_Click()
        timersettime = timersettime + 1
        If timersettime >= 60000 Then timersettime = 59999
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerSecMinus1_Click()
        timersettime = timersettime - 1
        If timersettime < 0 Then timersettime = 0
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdTimerSecMinus10_Click()
        timersettime = timersettime - 10
        If timersettime < 0 Then timersettime = 0
        timercountdowntime = timersettime
        Call TimerTimer_Timer
    End Sub
    Public Sub CmdLotteryTotalPlus10_Click()
        lotterytotal = lotterytotal + 10
        If lotterytotal >= 10000 Then lotterytotal = 9999
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryTotalPlus1_Click()
        lotterytotal = lotterytotal + 1
        If lotterytotal >= 10000 Then lotterytotal = 9999
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryTotalMinus1_Click()
        lotterytotal = lotterytotal - 1
        If lotterytotal < 2 Then lotterytotal = 2
        If lotterygroup > lotterytotal Then lotterytotal = lotterygroup
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryTotalMinus10_Click()
        lotterytotal = lotterytotal - 10
        If lotterytotal < 2 Then lotterytotal = 2
        If lotterygroup > lotterytotal Then lotterytotal = lotterygroup
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryGroupPlus10_Click()
        lotterygroup = lotterygroup + 10
        If lotterygroup >= 100 Then lotterygroup = 99
        If lotterygroup > lotterytotal Then lotterygroup = lotterytotal
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryGroupPlus1_Click()
        lotterygroup = lotterygroup + 1
        If lotterygroup >= 100 Then lotterygroup = 99
        If lotterygroup > lotterytotal Then lotterygroup = lotterytotal
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryGroupMinus1_Click()
        lotterygroup = lotterygroup - 1
        If lotterygroup < 2 Then lotterygroup = 2
        Call LotterySettingsRefresher
    End Sub
    Public Sub CmdLotteryGroupMinus10_Click()
        lotterygroup = lotterygroup - 10
        If lotterygroup < 2 Then lotterygroup = 2
        Call LotterySettingsRefresher
    End Sub
