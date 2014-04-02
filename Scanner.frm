VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Scanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "改良型嵌入式次世代苏打扫描仪"
   ClientHeight    =   1305
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4185
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2760
      Top             =   360
   End
   Begin VB.CheckBox Check1 
      Caption         =   "面神下凡模式"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7920
      Top             =   720
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Text            =   "230"
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "走着"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1931
      _Version        =   393216
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "172.16.20"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "000"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "0x7c921230"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Long
Dim dev As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer1.Interval = 30
Else
Timer1.Interval = 2000
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
If n = 255 Then
'n=1
If FrmAuth.n = "000" Then MsgBox "找不到苏打粉"
Unload Me



If Text1.Visible = True Then
List1.AddItem "-------- " & Text1.Text & ".x:" & Text2.Text & " --------"
Else
List1.AddItem ("end.")
End If



Command1.Enabled = True
Exit Sub
End If

ProgressBar1.Value = n
On Error Resume Next
Timer1.Enabled = True
Call Winsock1.Connect(Text1.Text & "." & n, Text2.Text)
End Sub

Private Sub Form_Load()
dev = False
End Sub

Private Sub Label1_DblClick()
Label1.Visible = False
Text1.Visible = True
Text2.Visible = True
Label2.Visible = True
End Sub

Private Sub List1_DblClick()
List1.Clear
End Sub

Private Sub ProgressBar1_Click()
Scanner.Height = 6795
dev = True
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 0 Then If Text1.Enabled = True Then List1.AddItem ("连接超了个时" & ":" & Text1.Text & "." & n)
Winsock1.Close
n = n + 1
Timer1.Enabled = False
Call Command1_Click
If n / 20 = n \ 20 Then Timer1.Interval = Timer1.Interval - 1
End Sub

Private Sub Timer2_Timer()
n = 1
Check1_Click
Timer2.Enabled = False

If dev = False Then Command1_Click
End Sub

Private Sub Winsock1_Connect()
Label2.Caption = n
If FrmAuth.n = "000" Then FrmAuth.n = n: FrmAuth.o = Text1.Text & "." & n: FrmAuth.p = Text2.Text
If Text1.Visible = True Then
List1.AddItem ("成功连接" & ":" & Text1.Text & "." & n & ":" & Text2.Text)

Else
List1.AddItem ("可以连接")
End If
Winsock1.Close
Timer1.Enabled = False
n = n + 1
Call Command1_Click
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
List1.AddItem (Description & ":" & Text1.Text & "." & n)
Winsock1.Close
Timer1.Enabled = False
n = n + 1
Call Command1_Click
End Sub
