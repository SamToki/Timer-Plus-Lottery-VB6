VERSION 5.00
Begin VB.Form FormInputNumber 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "Timer+Lottery"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
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
   Icon            =   "FormInputNumber.frx":0000
   LinkTopic       =   "FormInputNumber"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormInputNumber.frx":0CB2
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   5880
      Top             =   5565
   End
   Begin VB.Timer TimerCursorBlink 
      Interval        =   300
      Left            =   210
      Top             =   105
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Del"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3990
      MouseIcon       =   "FormInputNumber.frx":0E04
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4935
      Width           =   1905
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   315
      MouseIcon       =   "FormInputNumber.frx":0F56
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4935
      Width           =   1905
   End
   Begin VB.CommandButton CmdNumber0 
      Caption         =   "&0"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2415
      MouseIcon       =   "FormInputNumber.frx":10A8
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4725
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber9 
      Caption         =   "&9"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3780
      MouseIcon       =   "FormInputNumber.frx":11FA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3885
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber8 
      Caption         =   "&8"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2415
      MouseIcon       =   "FormInputNumber.frx":134C
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3885
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber7 
      Caption         =   "&7"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1050
      MouseIcon       =   "FormInputNumber.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3885
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber6 
      Caption         =   "&6"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3780
      MouseIcon       =   "FormInputNumber.frx":15F0
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3045
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber5 
      Caption         =   "&5"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2415
      MouseIcon       =   "FormInputNumber.frx":1742
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3045
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber4 
      Caption         =   "&4"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1050
      MouseIcon       =   "FormInputNumber.frx":1894
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3045
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber3 
      Caption         =   "&3"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3780
      MouseIcon       =   "FormInputNumber.frx":19E6
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2205
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber2 
      Caption         =   "&2"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2415
      MouseIcon       =   "FormInputNumber.frx":1B38
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2205
      Width           =   1380
   End
   Begin VB.CommandButton CmdNumber1 
      Caption         =   "&1"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1050
      MouseIcon       =   "FormInputNumber.frx":1C8A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2205
      Width           =   1380
   End
   Begin VB.Label LabelInputNumber4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   4725
      TabIndex        =   3
      Top             =   315
      Width           =   1170
   End
   Begin VB.Label LabelInputNumber3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   3255
      TabIndex        =   2
      Top             =   315
      Width           =   1170
   End
   Begin VB.Label LabelInputNumber2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1785
      TabIndex        =   1
      Top             =   315
      Width           =   1170
   End
   Begin VB.Label LabelHinttext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hold Alt key to enter number using keyboard"
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
      Height          =   330
      Left            =   315
      TabIndex        =   4
      Top             =   1700
      Width           =   5580
   End
   Begin VB.Label LabelInputNumber1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   1170
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   6210
   End
End
Attribute VB_Name = "FormInputNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

Private inputnumber1 As Integer
Private inputnumber2 As Integer
Private inputnumber3 As Integer
Private inputnumber4 As Integer

Public currentinputnumber As Integer

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    Public Sub TimerCursorBlink_Timer()
        Select Case currentinputnumber
            Case 1
                If LabelInputNumber1.Caption = ">" Then LabelInputNumber1.Caption = "" Else LabelInputNumber1.Caption = ">"
            Case 2
                If LabelInputNumber2.Caption = ">" Then LabelInputNumber2.Caption = "" Else LabelInputNumber2.Caption = ">"
            Case 3
                If LabelInputNumber3.Caption = ">" Then LabelInputNumber3.Caption = "" Else LabelInputNumber3.Caption = ">"
            Case 4
                If LabelInputNumber4.Caption = ">" Then LabelInputNumber4.Caption = "" Else LabelInputNumber4.Caption = ">"
        End Select
    End Sub

'[] COMMANDS []

    Public Sub CmdCancel_Click()
        inputnumber1 = 0
        inputnumber2 = 0
        inputnumber3 = 0
        inputnumber4 = 0
        currentinputnumber = 1
        LabelInputNumber1.Caption = ">"
        LabelInputNumber2.Caption = ">"
        LabelInputNumber3.Caption = ">"
        LabelInputNumber4.Caption = ">"
        FormMainWindow.Enabled = True: FormMiniMode.Enabled = True

        windowanimationtargetleft = (Screen.Width / 2)
        windowanimationtargettop = (Screen.Height / 2)
        windowanimationtargetwidth = 0
        windowanimationtargetheight = 0
    End Sub
    Public Sub CmdDelete_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 0
                inputnumber2 = 0
                inputnumber3 = 0
                inputnumber4 = 0
                LabelInputNumber1.Caption = ">"
                currentinputnumber = 1
            Case 2
                inputnumber1 = 0
                inputnumber2 = 0
                inputnumber3 = 0
                inputnumber4 = 0
                LabelInputNumber1.Caption = ">"
                LabelInputNumber2.Caption = ">"
                currentinputnumber = 1
            Case 3
                inputnumber2 = 0
                inputnumber3 = 0
                inputnumber4 = 0
                LabelInputNumber2.Caption = ">"
                LabelInputNumber3.Caption = ">"
                currentinputnumber = 2
            Case 4
                inputnumber3 = 0
                inputnumber4 = 0
                LabelInputNumber3.Caption = ">"
                LabelInputNumber4.Caption = ">"
                currentinputnumber = 3
        End Select
    End Sub
    
    Public Sub CmdNumber1_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 1
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 1
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 1
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 1
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber2_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 2
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 2
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 2
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 2
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber3_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 3
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 3
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 3
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 3
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber4_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 4
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 4
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 4
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 4
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber5_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 5
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 5
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 5
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 5
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber6_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 6
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 6
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 6
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 6
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber7_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 7
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 7
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 7
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 7
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber8_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 8
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 8
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 8
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 8
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber9_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 9
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 9
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 9
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 9
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub
    Public Sub CmdNumber0_Click()
        Select Case currentinputnumber
            Case 1
                inputnumber1 = 0
                LabelInputNumber1.Caption = inputnumber1
            Case 2
                inputnumber2 = 0
                LabelInputNumber2.Caption = inputnumber2
            Case 3
                inputnumber3 = 0
                LabelInputNumber3.Caption = inputnumber3
            Case 4
                inputnumber4 = 0
                LabelInputNumber4.Caption = inputnumber4
        End Select
        currentinputnumber = currentinputnumber + 1
        If currentinputnumber > FormMainWindow.inputnumberdigits Then Call InputNumberFinish: Exit Sub
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] SPECIAL []

    Public Sub InputNumberFinish()
        Select Case FormMainWindow.inputnumbermode
            Case "TimerMin"
                FormMainWindow.timersettimemin = inputnumber1 * 100 + inputnumber2 * 10 + inputnumber3
                FormMainWindow.timersettime = FormMainWindow.timersettimemin * 60 + FormMainWindow.timersettimesec
                FormMainWindow.timercountdowntime = FormMainWindow.timersettime
                Call FormMainWindow.TimerTimer_Timer
            Case "TimerSec"
                If (inputnumber1 * 10 + inputnumber2) > 59 Then
                    MsgBox "CAUTION: Number invalid." & vbCrLf & "[Second] cannot be larger than 59.", vbExclamation + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
                Else
                    FormMainWindow.timersettimesec = inputnumber1 * 10 + inputnumber2
                    FormMainWindow.timersettime = FormMainWindow.timersettimemin * 60 + FormMainWindow.timersettimesec
                    FormMainWindow.timercountdowntime = FormMainWindow.timersettime
                End If
                Call FormMainWindow.TimerTimer_Timer
            Case "LotteryTotal"
                If (inputnumber1 * 1000 + inputnumber2 * 100 + inputnumber3 * 10 + inputnumber4) < FormMainWindow.lotterygroup Then
                    MsgBox "CAUTION: Number invalid." & vbCrLf & "[Group] cannot be larger than the lottery number range.", vbExclamation + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
                Else
                    If (inputnumber1 * 1000 + inputnumber2 * 100 + inputnumber3 * 10 + inputnumber4) < 2 Then
                        MsgBox "The number cannot be smaller than 2.", vbExclamation + vbOKOnly + vbDefaultButton1, "Number Invalid"
                    Else
                        FormMainWindow.lotterytotal = inputnumber1 * 1000 + inputnumber2 * 100 + inputnumber3 * 10 + inputnumber4
                    End If
                End If
                Call FormMainWindow.LotterySettingsRefresher
            Case "LotteryGroup"
                If (inputnumber1 * 10 + inputnumber2) > FormMainWindow.lotterytotal Then
                    MsgBox "CAUTION: Number invalid." & vbCrLf & "[Group] cannot be larger than the lottery number range.", vbExclamation + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
                Else
                    If (inputnumber1 * 10 + inputnumber2) < 2 Then
                        MsgBox "CAUTION: Number invalid." & vbCrLf & "[Group] cannot be smaller than 2.", vbExclamation + vbOKOnly + vbDefaultButton1, "Timer+Lottery"
                    Else
                        FormMainWindow.lotterygroup = inputnumber1 * 10 + inputnumber2
                    End If
                End If
                Call FormMainWindow.LotterySettingsRefresher
        End Select

        Call CmdCancel_Click
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
