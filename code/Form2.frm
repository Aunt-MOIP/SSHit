VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4296
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Form1"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "D'tCM"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "debug"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hacking"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Donate 6CNY to order new features! "
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "本项目已全部由某跨国情报组织托管"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "for users"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "b Hackintosh or explode"
      Height          =   492
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   2772
   End
   Begin VB.Label Label5 
      Caption         =   "may cause fatal memory leaks"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "~ to the gate"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "MOIP 2012, be aware of cprt monster!"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Only for developers"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "More tweaks @synthesize"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim hackstr As String
hackstr = InputBox("0x7c921230", "check asm enter-point", "0x7c921234", Form2.Left, Form2.Top)
'If hackstr = "dksq%m#$ggzfdb" Then
If hackstr = "a" Then
hackstr = InputBox("String", "Hacking to the gate", "plink.exe 172.17.18.5 -N -ssh -2 -P 22 -l root -C -D 7070 -v -pw 5000", Form2.Left, Form2.Top)
changeCmdStr (hackstr)
End If
'Sleep 6000
End Sub

Private Sub Command2_Click()
Dim AA(500000) As String
For j = 1 To 500000
For i = 1 To 500000
Call Command2_Click
AA(i) = AA(i) & "FFFFFFFFFFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUUUUUCCCCCCCCCCCCCCCCCCCCKKKKKKKKKKKK"
AA(i) = AA(i) & "FFFFFFFFFFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUUUUUCCCCCCCCCCCCCCCCCCCCKKKKKKKKKKKK"
AA(i) = AA(i) & "FFFFFFFFFFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUUUUUCCCCCCCCCCCCCCCCCCCCKKKKKKKKKKKK"
AA(i) = AA(i) & "FFFFFFFFFFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUUUUUCCCCCCCCCCCCCCCCCCCCKKKKKKKKKKKK"
AA(i) = AA(i) & "FFFFFFFFFFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUUUUUCCCCCCCCCCCCCCCCCCCCKKKKKKKKKKKK"
Next i
Next j


End Sub

Private Sub Command3_Click()
frmHide.Top = 0
frmHide.Left = 0
frmHide.Height = Screen.Height
frmHide.Width = Screen.Width
frmHide.Visible = True
Sleep 600000
End Sub

Private Sub Command4_Click()
Form1.Top = Form2.Top
Form1.Left = Form2.Left
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub Form_Load()
Form2.Top = Form1.Top
Form2.Left = Form1.Left
End Sub

