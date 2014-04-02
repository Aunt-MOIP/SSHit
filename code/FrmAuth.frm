VERSION 5.00
Begin VB.Form FrmAuth 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FrmAuth"
   ClientHeight    =   4080
   ClientLeft      =   15900
   ClientTop       =   9360
   ClientWidth     =   6660
   LinkTopic       =   "Form6"
   ScaleHeight     =   4080
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrL2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Timer tmrL1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1440
      Top             =   2520
   End
   Begin VB.Timer tmrAbsTime 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   2880
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5400
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   5400
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label p 
      Height          =   252
      Left            =   5040
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label o 
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label n 
      Caption         =   "000"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4320
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      Caption         =   "长相薄喜欢楠神"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "苏打扫描仪"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "天王盖地虎"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1680
   End
End
Attribute VB_Name = "FrmAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const aaa& = -1
Private Const b& = &H1
Private Const c& = &H2
Dim bmwPassword As String
Dim a, X As Double
Dim fact As Double
Dim tpalpha As Long
Dim absTime As Double
Dim cll As Integer
Dim fuckCounter As Long
Private Sub Command1_Click()
ads = Noti("今天天不错", "张搏翔翔翔翔翔翔翔翔翔翔翔翔翔翔", vbWhite, 600, False)
Timer3.Enabled = True
End Sub

Private Sub Form_Click()

fuckCounter = fuckCounter + 1
If fuckCounter > 20 Then
If Timer1.Enabled = True Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End If
End Sub

Private Sub Form_DblClick()
fuckCounter = fuckCounter + 2
If fuckCounter > 20 Then
If Timer1.Enabled = True Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End If
Command1.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 23 Then
fact = 30
End If

End Sub

Private Sub Form_Load()
fact = 15
fuckCounter = 0
If VBA.Environ("OS") = "Windows_NT" Then
Timer1.Interval = 1
'nil = Noti("XP爱好者", "Windows XP爱好者！", vbCyan, 15000, False)
End If
Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, 125, LWA_ALPHA
    tpalpha = 125
a = 0
s = 1000000
SetWindowPos Me.hWnd, aaa, 0, 0, 0, 0, b Or c

Timer1.Enabled = True
Me.Visible = True


End Sub

Private Sub Form_Resize()
Text1.Left = (Me.ScaleWidth - Text1.Width) / 2
Text1.Top = Me.Height / 2 - 2 * Text1.Height
Label1.Left = Text1.Left
Label2.Left = Label1.Left + 1800
Label3.Left = Text1.Left + Text1.Width
Label1.Top = Text1.Top - 1560
Label2.Top = Label1.Top
Label3.Top = Text1.Top
End Sub

Private Sub Label1_Click()
    changeCmdStr "plink.exe 172.16.20.56  -N -ssh -2 -P 22 -l root -C -D 0.0.0.0:7070 -v -pw alpine"
    Timer4.Enabled = True
    'Label1.BackColor = &H8080FF
    Label2.Visible = False
    
    'Label1.Left = (Me.Width - Label1.Width) / 2
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HFF
Label1.Caption = "长按隐身模式"
'ni = Noti("Hold to stealth mode", "", RGB(100, 100, 100), 100, False)
tmrL1.Enabled = True
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HC0&
Label1.Caption = "天王盖地虎"
tmrL1.Enabled = False
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF00&
Label2.Caption = "长按隐身模式"
'ni = Noti("Hold to stealth mode", "", RGB(100, 100, 100), 100, False)
tmrL2.Enabled = True
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC000&
Label2.Caption = "宝塔镇河妖"
tmrL2.Enabled = False
End Sub

Private Sub Label2_Click()
Me.Visible = False
Scanner.Show 1
Unload Scanner
Me.Visible = True
    If n <> "000" Then
    changeCmdStr "plink.exe " & FrmAuth.o & "  -N -ssh -2 -P " & p & " -l root -C -D 0.0.0.0:7070 -v -pw alpine"
    Timer4.Enabled = True
    Label1.Visible = False
         'Label2.Left = (Me.Width - Label2.Width) / 2
End If
End Sub




Private Sub Label3_DblClick()

i = Len(Text1.Text)
Text1.Text = Text1.Text + Mid("adenosinetriphosphate", i + 1, 1)
If Text1.Text = "adenosinetriphosphate" Then

Call Text1_KeyPress(13)
ni = Noti("努力达人！", "密码是adenosinetriphosphate一定要记住", vbRed, 30000, False)
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If tmrAbsTime.Enabled = False Then tmrAbsTime.Enabled = True
If KeyAscii = 13 Then
If Text1.Text = "adenosinetriphosphate" Or Text1.Text = "" Then
Call Auth_Succeed
If absTime < 0.5 And Text1.Text = "might" Then ni = Noti("手速爱好者！", "在" & absTime & "秒内成功输入密码", vbRed, 2000, False)
If absTime < 1 And Text1.Text = "adenosinetriphosphate" Then ni = Noti("莫作弊", "阿姨最不喜欢小伙纸作弊了", vbWhite, 10000, False)
If absTime < 3 And absTime > 1 And Text1.Text = "adenosinetriphosphate" Then ni = Noti("三磷酸腺苷！", "在" & absTime & "秒内输入如此长的密码", vbWhite, 2000, False)
If absTime < 1.73 And absTime > 1 And Text1.Text = "adenosinetriphosphate" Then ni = Noti("世界纪录！", "" & absTime & "秒打破了阿姨创造的世界纪录", vbWhite, 2000, False)

If Text1.Text = "adenosinetriphosphate" Then ni = Noti("汪汪汪汪汪汪！", "原来是知情者汪先生，失礼！", vbYellow, 10000, False)
If Hour(Now) >= 18 And Hour(Now) < 19 Then ni = Noti("同学静校了！", "6点以后还在上网", vbCyan, 15000, False)
If Hour(Now) >= 19 And Hour(Now) < 24 Then ni = Noti("翻墙爱好者！", "翻墙有风险，调戏保安需谨慎", vbCyan, 15000, False)


Else
If Text1.Text = "fuck" Or Text1.Text = "cao" Or Text1.Text = "sb" Or Text1.Text = "shabi" Then
ni = Noti("滚！", "呵呵", vbWhite, 2000, False)
End If
If Text1.Text = "" Then
Me.BackColor = vbRed
Timer3.Enabled = True
Exit Sub
End If


ni = Noti("智商捉鸡！", "密码又错了", vbGreen, 100, False)
Me.BackColor = vbRed
Timer3.Enabled = True
Exit Sub

End If
End If
End Sub

Private Sub Timer1_Timer()
a = a + 1
X = X + fact * a
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Width = Abs(X)
Me.Height = Abs(X)
If X > 2 * Screen.Height Then
Timer1.Enabled = False

Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If tpalpha < 230 Then
tpalpha = tpalpha + 2
  Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, tpalpha, LWA_ALPHA
    
    Else
    Timer2.Enabled = False
    Text1.Visible = True
    Label3.Visible = True
    
End If
End Sub

Private Sub Timer3_Timer()
If tpalpha > 10 Then
tpalpha = tpalpha - 5
  Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, tpalpha, LWA_ALPHA
    
    Else
    Timer3.Enabled = False
    End
    Text1.Visible = False
    Label3.Visible = False
    Unload Me
    Form1.Show
End If
End Sub

Private Sub Timer4_Timer()
If tpalpha > 5 Then
tpalpha = tpalpha - 5
  Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, tpalpha, LWA_ALPHA
    
    Else
Timer4.Enabled = False
Text1.Visible = False
Label3.Visible = False
Unload Me
Form1.Show
End If
End Sub

Sub Auth_Succeed()
    Text1.Visible = False
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = False

End Sub

Private Sub tmrAbsTime_Timer()
absTime = absTime + 0.01
End Sub

Private Sub tmrL1_Timer()
nil = Silent
Call Label1_Click
tmrL1.Enabled = False
End Sub

Private Sub tmrL2_Timer()
nil = Silent
Call Label2_Click
tmrL2.Enabled = False
End Sub
