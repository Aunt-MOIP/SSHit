Attribute VB_Name = "ModMain"
Dim Q(10) As Long
Dim isSilent As Boolean
Sub Main()


End Sub



Public Function Noti(Caption As String, Content As String, Color As ColorConstants, Delay As Long, isBug As Boolean)
 If isSilent = True Then Exit Function
Dim a As New Notify
Q(0) = Q(0) + 1
For i = 1 To 9
If Q(i) = 0 Then
Q(i) = 1
Noti = i
a.Top = (100 + a.Height) * i - a.Height

a.txtDelay = Delay / 100
a.lblTitle.Caption = Caption
a.lblContent.Caption = Content
a.txtQueue = i
a.txtIsBug = isBug
a.BackColor = Color
a.lblContent.BackColor = Color
a.lblTitle.BackColor = Color
a.Visible = True
Exit For
End If
Next i

End Function

Public Function NotiFinal(QQ As Long)
Q(QQ) = 0

End Function

Public Function Silent()
ni = Noti("隐身模式！", "保证不会被自己发现哦～", vbWhite, 100000, False)
isSilent = True
End Function
Public Function noSilent()
isSilent = False
End Function
