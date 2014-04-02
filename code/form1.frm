VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "好无聊啊"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   3360
   End
   Begin VB.CommandButton Command7 
      Caption         =   "c"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3960
      Top             =   1440
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Form2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer refreshCLI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   960
   End
   Begin VB.Label Label8 
      Caption         =   "本项目已全部由某跨国情报组织托管"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "关于"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "for developers"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   2652
   End
   Begin VB.Label Label5 
      Caption         =   "隐藏"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "连接网络..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "MOIP 2012, be aware of cprt monster!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Hacking, Coding and GUI: MOIP"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to SSHit 7779 - It is Beta!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cliBuff As String
Dim iackBuffer As String
Dim inputBuffer As String
Dim StrExe As String
Dim isConnOK As Boolean
Dim isFirstHide As Boolean
Dim ii As Integer
Dim isZhangxiangboShabi As Boolean
Dim retryTimes As Long
Dim duanjisong As Long
Dim skqDiaobaole As Boolean

Sub Cmdexe_Click()
 Dim Ret As Long
 
'StrExe
 Ret = DosInput(StrExe)
 
 If Ret <> 0 Then
   
   'MsgBox "在写入控制台管道的时候出现错误", vbInformation, "错误"
   Exit Sub
 End If
 Print """" & StrExe & """"
 StrExe = ""
End Sub

Private Sub CmdGet()
  Dim strR As String
 Ret = DosOutput(strR)
 If Ret = 0 Then
  TxtOutput.Text = strR
 Else
 Form1.Caption = "读取控制台输出错误"
 ni = Noti("又见Bug！", "又出bug了，重启，重启一下", vbRed, 10000, False)
 End If
 'sw True
End Sub




Private Sub Command1_Click()

Form1.Visible = False
Form3.Top = Form1.Top
Form3.Left = Form1.Left
Form3.Show
End Sub

Private Sub Command2_Click()
If isFirstHide Then
MsgBox "双击屏幕右上角恢复", vbSystemModal
isFirstHide = False
End If
Form1.Visible = False
frmHide.Visible = True
End Sub

Private Sub Command3_Click()
Form1.Visible = False
Form2.Visible = True
Form2.Top = Form1.Top
Form2.Left = Form1.Left


End Sub

Private Sub Command4_Click()
Form4.Show 1
End Sub

Private Sub Command5_Click()
isFirstHide = False
Form1.Visible = False
frmHide.Visible = True
End Sub

Private Sub Command6_Click()
EndDosIo
End
End Sub

Private Sub Command7_Click()
Dim passs As String

passs = InputBox("某跨国情报组织扫地大队要求身份", "鉴定安全密码", , Form1.Left, Form1.Top)

If passs = "1394" Then

Form5.Show

Else

MsgBox "BIG BROTHER IS WATCHING YOU", vbCritical, "FORBIDDEN"

End If

End Sub

Private Sub Command8_Click()
duanjisong = duanjisong + 1
If duanjisong <= 6 Then
ni = Noti("目击整个过程的刘先生", "我活了42年见过的最无聊的人", vbCyan, 8000, False)
Else
 ni = Noti("目击整个过程的Z.X.B.", "我活了17年见过的最无聊的人", vbRed, 18000, False)
End If
End Sub

Private Sub Form_DblClick()
'MsgBox "我只是无数Form1中的一个", vbSystemModal
ni = Noti("我只是无数Form1中的一个!", "打开real debug模式！", vbRed, 8000, False)
Command7.Visible = True
End Sub

Private Sub Form_Loadthen()
 Dim Ret As Long
 a = Shell("taskkill /f /im plink.exe", vbHide)
 'a = Shell("regedit /s PROXY.reg")
 Sleep 1000
 Ret = InitDosIO()
 
 If Ret <> 0 Then
   MsgBox "没有找到plink" & vbCrLf & vbCrLf & "确认下本目录里有plink.Exe" & "——阿姨温馨小提示"
    ni = Noti("又见Bug！", "确认下本目录里有plink.Exe", vbRed, 10000, False)

   Exit Sub
 End If
 
 
Command1.Enabled = False
isFirstHide = True
Timer1.Enabled = True
Timer2.Enabled = False

Timer4.Enabled = True
refreshCLI.Enabled = True
End Sub

Private Sub Form_Load()

Call Form_Loadthen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
EndDosIo
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndDosIo
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub refreshCLI_Timer() '
On Error Resume Next
  Dim strR As String
  Ret = DosOutput(strR)
  'cliBuff = cliBuff + strR
  Form5.Text1.SelStart = Len(Form5.Text1)
  Form5.Text1.SelText = Form5.Text1.SelText + strR
  cliBuff = Form5.Text1
  If Ret = 0 And Form5.Text1.SelText <> "" And isConnOK = False Then
        Form5.Text1.SelText = Replace(Form5.Text1.SelText, "172.17.18", "127.0.0")
        Form5.Text1.SelText = Replace(Form5.Text1.SelText, "1165", "65536")
        
      
 
  End If

     '------------------------
     If Replace(cliBuff, "password", "") <> mon1 Then
      If Replace(cliBuff, "Sent", "") <> mon1 Then
      Else
        StrExe = "alpine"
        Call Cmdexe_Click
        End If
     End If
     
     
     '-------------OverRideALL Y
     'Store key in cache? (y/n)
     If Replace(cliBuff, "y/n", "") <> cliBuff And Replace(cliBuff, "dynamic", "") = cliBuff Then
      
        StrExe = "y"
        Call Cmdexe_Click

     End If
     '-------------/OverRideALL Y
     
     
     '--conn ok
     If isConnOK = False And Replace(cliBuff, "dynamic forwarding", "") <> cliBuff Then
        
        Call connOK
        
     End If
     
     '----
     If Replace(cliBuff, "timeout", "") <> cliBuff Then
           
        Call connFail
    
     End If
    If Replace(cliBuff, "FATAL", "") <> cliBuff Then
         
        Call connFail
    
    End If
    
    If Replace(cliBuff, "error", "") <> cliBuff Then
         
        Call connFail
    
    End If
    'denied
    
    If Replace(cliBuff, "denied", "") <> cliBuff Then
         refreshCLI.Enabled = False
        Form1.Caption = "Form1 - 已过期"
        
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
    'MsgBox "安全策略已更新,无法提供安全连接,请请求更新的版本.", vbSystemModal
    ni = Noti("你OUT了！", "安全策略已更新,无法提供安全连接,请请求更新的版本.", vbRed, 15000, False)
    End If
    'Form5.Text1 = cliBuff
    Form5.Caption = Len(cliBuff)
 'sw True
' Form1.Caption = Timer1.Interval
End Sub

Sub connOK()
Command1.Enabled = True
isConnOK = True
Print ("OK")
Form1.Caption = "Form1 - [Active]"
Timer2.Enabled = True
If skqDiaobaole = False Then
ni = Noti("接上了！", "连接上网络！", vbGreen, 4000, False)
skqDiaobaole = True
End If
End Sub

Sub connFail()
isConnOK = False
Print ("Fail")
Form5.Text1.Text = ""
retryTimes = retryTimes + 1
If retryTimes > 5 Then
  ni = Noti("出Bug了！", "重启，重启一下。。。", vbWhite, 5000, False)
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
refreshCLI.Enabled = False
EndDosIo

End If

 cliBuff = ""
iackBuffer = ""
inputBuffer = ""
 StrExe = ""
 isConnOK = False
 
 EndDosIo
  Ret = InitDosIO()
 
 If Ret <> 0 Then
   MsgBox "没有找到plink" & vbCrLf & vbCrLf & "确认下本目录里有plink.Exe" & "——阿姨温馨小提示"
   Exit Sub
 End If
 

End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
If isConnOK = False Then
  ni = Noti("台母熬特！", "连。。。接。。。超。。。时。。。", vbRed, 10000, False)

End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Print "Ready"
End Sub


Private Sub Timer4_Timer()
If isZhangxiangboShabi = True Then
Form1.Visible = False
frmHide.Visible = True
End If

End Sub
