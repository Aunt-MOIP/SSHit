VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   6984
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6468
   LinkTopic       =   "Form5"
   ScaleHeight     =   6984
   ScaleWidth      =   6468
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "°²ÖäÃ¿ÑêÏ²»¶Ê¢ÖÇ²©"
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "³¤Ïà±¡Ï²»¶éªÉñ"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Form5.Hide
End Sub

