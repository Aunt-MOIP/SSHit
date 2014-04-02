VERSION 5.00
Begin VB.Form Notify 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   LinkTopic       =   "Form6"
   ScaleHeight     =   1230
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Left            =   1320
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   600
   End
   Begin VB.TextBox txtIsBug 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4200
      TabIndex        =   4
      Text            =   "0"
      Top             =   -480
      Width           =   615
   End
   Begin VB.TextBox txtDelay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4560
      TabIndex        =   3
      Text            =   "0"
      Top             =   -360
      Width           =   615
   End
   Begin VB.TextBox txtQueue 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4560
      TabIndex        =   0
      Text            =   "0"
      Top             =   -120
      Width           =   615
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lblContent"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lblTitle"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Notify"
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
Dim NotiAlp As Long
Dim initAlpha As Long
Private Sub Form_Load()
Randomize
initAlpha = 205
Me.Left = Screen.Width - Me.Width - 400

Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, initAlpha, LWA_ALPHA
    NotiAlp = initAlpha
    
    
SetWindowPos Me.hWnd, aaa, 0, 0, 0, 0, b Or c

End Sub

Private Sub Text1_Change()

End Sub

Private Sub lblTitle_DblClick()
NotiAlp = 20
End Sub

Private Sub tmrDelay_Timer()
'SetWindowPos Me.hWnd, aaa, 0, 0, 0, 0, b Or c

NotiAlp = NotiAlp - 2
Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, NotiAlp, LWA_ALPHA

If txtIsBug = True Then
Me.Top = Me.Top + Fix(75 * Rnd())
Me.Left = Me.Left - Fix(100 * Rnd())
End If

    
If NotiAlp <= initAlpha - 200 Then
a = NotiFinal(txtQueue)
Unload Me
End If
End Sub

Private Sub txtDelay_Change()
tmrDelay.Interval = txtDelay.Text
tmrDelay.Enabled = True

End Sub
