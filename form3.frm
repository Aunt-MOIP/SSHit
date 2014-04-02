VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "NeXT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "本文档将指导使用本程序连接到Internet."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2

Dim Flip As Integer
Private Sub Command1_Click()
Flip = Flip + 1
Select Case Flip
Case 1
Label1.Caption = "打开Internet Explorer"
Case 2
Label1.Caption = "菜单栏 工具 - Internet 选项"
Case 3
Label1.Caption = "连接选项卡 - 局域网设置"
Case 4
Label1.Caption = "勾选 为 LAN 使用代理服务器"
Case 5
Label1.Caption = "点击 高级 按钮，取消勾选 对所有协议均使用相同的代理服务器"
Case 6
Label1.Caption = "仅在套接字/Socks (C)一行输入" & vbCrLf & vbCrLf & " 127.0.0.1:7070"
Case 7
Label1.Caption = "确定 - 确定 - 确定" & vbCrLf & "Enjoy the Internet!"
Command1.Caption = "返回"
Case 8
Form1.Left = Form3.Left
Form1.Top = Form3.Top
Form1.Visible = True
Unload Me
End Select

End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c

Flip = 0
Command1.Caption = "NeXT"
End Sub

