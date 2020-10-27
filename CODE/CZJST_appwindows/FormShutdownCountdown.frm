VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormShutdownCountdown 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "Timer+Lottery"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12510
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
   Icon            =   "FormShutdownCountdown.frx":0000
   LinkTopic       =   "FormShutdownCountdown"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormShutdownCountdown.frx":0CB2
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   12180
      Top             =   2415
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Height          =   645
      Left            =   1995
      MouseIcon       =   "FormShutdownCountdown.frx":0E04
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1890
      Width           =   5055
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Shut Down Now (30)"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7245
      MouseIcon       =   "FormShutdownCountdown.frx":0F56
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1890
      Width           =   5055
   End
   Begin VB.Timer TimerShutdownCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7350
      Top             =   2415
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   420
      Left            =   7875
      TabIndex        =   5
      Top             =   2415
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
   Begin VB.Label LabelHinttextB 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "HintTextB Abg"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   315
      TabIndex        =   2
      Top             =   1100
      Width           =   11880
   End
   Begin VB.Label LabelAppTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Timer+Lottery"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   160
      TabIndex        =   0
      Top             =   105
      Width           =   10005
   End
   Begin VB.Label LabelHinttextA 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "HintTextA Abg"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   315
      TabIndex        =   1
      Top             =   500
      Width           =   11880
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   2745
      Left            =   0
      Top             =   0
      Width           =   12510
   End
End
Attribute VB_Name = "FormShutdownCountdown"
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

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Public Sub Form_Load()
        'ALWAYS FRONT (CODES FROM INTERNET)
        retValue = SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 1, 1, SWP_SHOWWINDOW)
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    Public Sub TimerShutdownCountdown_Timer()
        FormMainWindow.shutdowncountdowntimeout = FormMainWindow.shutdowncountdowntimeout - 1
        Select Case FormMainWindow.shutdowncountdowntype
            Case "Shutdown"
                CmdOK.Caption = "Shut Down Now (" & FormMainWindow.shutdowncountdowntimeout & ")"
            Case "Restart"
                CmdOK.Caption = "Restart Now (" & FormMainWindow.shutdowncountdowntimeout & ")"
        End Select
        If FormMainWindow.shutdowncountdowntimeout <= 0 Then Call CmdOK_Click
    End Sub

'[] COMMANDS []

    Public Sub CmdCancel_Click()
        TimerShutdownCountdown.Enabled = False
        FormMainWindow.Enabled = True: FormMiniMode.Enabled = True

        windowanimationtargetleft = (Screen.Width / 2) - (12510 / 2)
        windowanimationtargettop = 0
        windowanimationtargetwidth = 12510
        windowanimationtargetheight = 0
    End Sub
    Public Sub CmdOK_Click()
        'Interface sound...
        If FormMainWindow.soundswitch = True Then
            Select Case FormMainWindow.interfacesoundswitch
                Case True
                    WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Shutdown.wav"
                Case False
                    WindowsMediaPlayer1.URL = ""
            End Select
        End If

        Select Case FormMainWindow.shutdowncountdowntype
            Case "Shutdown"
                LabelHinttextA.Caption = "Shutting down..."
                Shell "cmd.exe /c shutdown -s -t 0", vbHide
            Case "Restart"
                LabelHinttextA.Caption = "Restarting computer..."
                Shell "cmd.exe /c shutdown -r -t 0", vbHide
        End Select

        CmdCancel.Visible = False
        CmdOK.Visible = False
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerWindowAnimation_Timer()
        If ((Me.Left = windowanimationtargetleft) And (Me.Top = windowanimationtargettop) And (Me.Width = windowanimationtargetwidth) And (Me.Height = windowanimationtargetheight)) Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If Me.Left > windowanimationtargetleft Then Me.Left = Me.Left - Abs(Me.Left - windowanimationtargetleft) / 8  'This case must be slower than others...
                If Me.Left < windowanimationtargetleft Then Me.Left = Me.Left + Abs(Me.Left - windowanimationtargetleft) / 8
                If Me.Top > windowanimationtargettop Then Me.Top = Me.Top - Abs(Me.Top - windowanimationtargettop) / 8
                If Me.Top < windowanimationtargettop Then Me.Top = Me.Top + Abs(Me.Top - windowanimationtargettop) / 8
                If Me.Width > windowanimationtargetwidth Then Me.Width = Me.Width - Abs(Me.Width - windowanimationtargetwidth) / 8
                If Me.Width < windowanimationtargetwidth Then Me.Width = Me.Width + Abs(Me.Width - windowanimationtargetwidth) / 8
                If Me.Height > windowanimationtargetheight Then Me.Height = Me.Height - Abs(Me.Height - windowanimationtargetheight) / 8
                If Me.Height < windowanimationtargetheight Then Me.Height = Me.Height + Abs(Me.Height - windowanimationtargetheight) / 8
                If Abs(Me.Left - windowanimationtargetleft) < 10 Then Me.Left = windowanimationtargetleft
                If Abs(Me.Top - windowanimationtargettop) < 10 Then Me.Top = windowanimationtargettop
                If Abs(Me.Width - windowanimationtargetwidth) < 10 Then Me.Width = windowanimationtargetwidth
                If Abs(Me.Height - windowanimationtargetheight) < 10 Then Me.Height = windowanimationtargetheight
            Case False
                Me.Move windowanimationtargetleft, windowanimationtargettop, windowanimationtargetwidth, windowanimationtargetheight
        End Select

        If windowanimationtargetheight = 0 And Me.Height < 100 Then Me.Hide  'This case must be slower than others...
    End Sub
