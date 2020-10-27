VERSION 5.00
Begin VB.Form FormLottery 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Timer+Lottery"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
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
   Icon            =   "FormLottery.frx":0000
   LinkTopic       =   "FormLottery"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormLottery.frx":0CB2
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10395
      MouseIcon       =   "FormLottery.frx":0E04
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2730
      Width           =   4530
   End
   Begin VB.Timer TimerScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13860
      Top             =   105
   End
   Begin VB.Timer TimerLottery 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   14385
      Top             =   105
   End
   Begin VB.PictureBox PictureboxScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2475
      Left            =   105
      ScaleHeight     =   2558.544
      ScaleMode       =   0  'User
      ScaleWidth      =   15150
      TabIndex        =   1
      Top             =   105
      Width           =   15150
      Begin VB.Label LabelScrollText 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ScrText Abg"
         BeginProperty Font 
            Name            =   "Avenir Next LT Pro"
            Size            =   111.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2775
         Left            =   105
         TabIndex        =   2
         Top             =   -290
         Width           =   14000
      End
   End
   Begin VB.Label LabelHintText 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "HintText Abg"
      BeginProperty Font 
         Name            =   "Avenir Next LT Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF6600&
      Height          =   1230
      Left            =   210
      TabIndex        =   3
      Top             =   2400
      Width           =   9900
   End
   Begin VB.Label LabelAppTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Timer+Lottery"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   3660
      Width           =   8010
   End
End
Attribute VB_Name = "FormLottery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

    'ALWAYS FRONT (CODES FROM INTERNET)
        'DISABLED: Dim retValue As Long
        'DISABLED: Private Declare Function SetWindowPos Lib "user32" ( _
            'DISABLED: ByVal hWnd As Long, _
            'DISABLED: ByVal hWndInsertAfter As Long, _
            'DISABLED: ByVal x As Long, ByVal y As Long, _
            'DISABLED: ByVal cX As Long, ByVal cY As Long, _
            'DISABLED: ByVal wFlags As Long _
            'DISABLED: ) As Long
            'DISABLED: Const HWND_TOPMOST = -1
            'DISABLED: Const SWP_SHOWWINDOW = &H40

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
        'DISABLED: retValue = SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 1, 1, SWP_SHOWWINDOW)
        'HALF TRANSPARENT (CODES FROM INTERNET)
        MakeTransparent Me.hWnd, 0

        TimerScroll.Interval = 40
        TimerLottery.Interval = 50

        PictureboxScroll.Left = 20000
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    Public Sub TimerScroll_Timer()
        If FormMainWindow.lotteryscrollprogress = 0 Then Call Form_Load

        If FormMainWindow.lotteryscrollprogress = 35 Then
            FormMainWindow.lotterynumberrecordX = FormMainWindow.lotterynumberrecord9
            FormMainWindow.lotterynumberrecord9 = FormMainWindow.lotterynumberrecord8
            FormMainWindow.lotterynumberrecord8 = FormMainWindow.lotterynumberrecord7
            FormMainWindow.lotterynumberrecord7 = FormMainWindow.lotterynumberrecord6
            FormMainWindow.lotterynumberrecord6 = FormMainWindow.lotterynumberrecord5
            FormMainWindow.lotterynumberrecord5 = FormMainWindow.lotterynumberrecord4
            FormMainWindow.lotterynumberrecord4 = FormMainWindow.lotterynumberrecord3
            FormMainWindow.lotterynumberrecord3 = FormMainWindow.lotterynumberrecord2
            FormMainWindow.lotterynumberrecord2 = FormMainWindow.lotterynumberrecord1
            FormMainWindow.lotterynumberrecord1 = FormMainWindow.lotterynumber
            Call FormMainWindow.LotteryRecordsRefresher
            CmdCancel.Enabled = False
        End If

        If FormMainWindow.lotteryscrollprogress = 60 Then
            FormMainWindow.lotteryscrollprogress = 61
            TimerLottery.Enabled = False
            TimerScroll.Enabled = False
            FormMainWindow.ShapeLightLottery.BorderStyle = 0
            FormMainWindow.ShapeLightLottery.FillStyle = 1
            FormMiniMode.ShapeLightLottery.BorderStyle = 0
            FormMiniMode.ShapeLightLottery.FillStyle = 1
            FormMainWindow.Enabled = True: FormMiniMode.Enabled = True: FormLottery.Hide
            Exit Sub
        End If

        FormMainWindow.lotteryscrollprogress = FormMainWindow.lotteryscrollprogress + 1

        Select Case FormMainWindow.useoldscrollanimationinlotterywindowswitch
            Case False
                'TOTAL LENGTH: 20000
                Select Case FormMainWindow.lotteryscrollprogress
                    Case 1 To 10
                        PictureboxScroll.Left = PictureboxScroll.Left - (3714 - 348 * (FormMainWindow.lotteryscrollprogress - 0))
                        If FormMainWindow.setanimationswitch = True Then
                            MakeTransparent Me.hWnd, Int(24 * FormMainWindow.lotteryscrollprogress)
                        Else
                            MakeTransparent Me.hWnd, 240
                        End If
                    Case 11 To 20
                        PictureboxScroll.Left = PictureboxScroll.Left - (61 - 2 * (FormMainWindow.lotteryscrollprogress - 10))
                        MakeTransparent Me.hWnd, 240
                    Case 21 To 30
                        PictureboxScroll.Left = PictureboxScroll.Left - (41 - 2 * (FormMainWindow.lotteryscrollprogress - 20))
                        MakeTransparent Me.hWnd, 240
                    Case 31 To 40
                        PictureboxScroll.Left = PictureboxScroll.Left - (20.5 - 1 * (FormMainWindow.lotteryscrollprogress - 30))
                        MakeTransparent Me.hWnd, 240
                    Case 41 To 50
                        PictureboxScroll.Left = PictureboxScroll.Left - (10.5 - 1 * (FormMainWindow.lotteryscrollprogress - 40))
                        MakeTransparent Me.hWnd, 240
                    Case 51 To 60
                        PictureboxScroll.Left = PictureboxScroll.Left - (-100 + 200 * (FormMainWindow.lotteryscrollprogress - 50))
                        If FormMainWindow.setanimationswitch = True Then
                            MakeTransparent Me.hWnd, Int(24 * (60 - FormMainWindow.lotteryscrollprogress))
                        Else
                            MakeTransparent Me.hWnd, 240
                        End If
                End Select
            Case True
                'TOTAL LENGTH: 20000
                Select Case FormMainWindow.lotteryscrollprogress
                    Case 1 To 10
                        PictureboxScroll.Left = PictureboxScroll.Left - 1800
                        If FormMainWindow.setanimationswitch = True Then
                            MakeTransparent Me.hWnd, Int(24 * FormMainWindow.lotteryscrollprogress)
                        Else
                            MakeTransparent Me.hWnd, 240
                        End If
                    Case 11 To 50
                        PictureboxScroll.Left = PictureboxScroll.Left - 25
                        MakeTransparent Me.hWnd, 240
                    Case 51 To 60
                        PictureboxScroll.Left = PictureboxScroll.Left - 1000
                        If FormMainWindow.setanimationswitch = True Then
                            MakeTransparent Me.hWnd, Int(24 * (60 - FormMainWindow.lotteryscrollprogress))
                        Else
                            MakeTransparent Me.hWnd, 240
                        End If
                End Select
        End Select
    End Sub

    Public Sub TimerLottery_Timer()
        If FormMainWindow.lotteryscrollprogress <= 20 Then
            Call FormMainWindow.RandomNumberGenerator
            Select Case FormMainWindow.lotterygroupswitch
                Case False
                    LabelScrollText.Caption = FormMainWindow.lotterynumber + 1
                Case True
                    LabelScrollText.Caption = (Int(FormMainWindow.lotterynumber / Int(FormMainWindow.lotterytotal / FormMainWindow.lotterygroup)) + 1) & " - " & (FormMainWindow.lotterynumber Mod Int(FormMainWindow.lotterytotal / FormMainWindow.lotterygroup) + 1)
            End Select
        End If

        If FormMainWindow.lotteryscrollprogress > 20 And FormMainWindow.lotteryscrollprogress < 35 Then
            'ANTI-REPEAT
            If (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord1) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord2) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord3) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord4) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord5) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord6) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord7) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord8) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecord9) And (FormMainWindow.lotterynumber <> FormMainWindow.lotterynumberrecordX) Then
                LabelHintText.Caption = "OK!"
            Else
                LabelHintText.Caption = "Duplicate item found!"
                If FormMainWindow.lotterypreventrepeatswitch = True Then
                    LabelHintText.Caption = "Duplicate item found!"
                    If FormMainWindow.lotteryscrollprogress > 31 Then
                        FormMainWindow.lotteryscrollprogress = 0
                        PictureboxScroll.Left = 20000
                        LabelHintText.Caption = "Retrying..."
                    End If
                End If
            End If
        End If
    End Sub

'[] COMMANDS []

    Public Sub CmdCancel_Click()
        FormMainWindow.WindowsMediaPlayer1.URL = ""
        TimerLottery.Enabled = False
        LabelScrollText.Caption = "0"
        LabelHintText.Caption = "Cancelled!"
        FormMainWindow.lotteryscrollprogress = 36
        CmdCancel.Enabled = False
    End Sub
