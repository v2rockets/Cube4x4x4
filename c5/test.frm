VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim G(), VGo(), WGo(), StoneLay(), Turn, cL(64, 7), nL(0 To 76, -1 To 1), sL(0 To 76, 4) As Integer
Dim p, l, Depth, Stone, n, Suggest, ToGo, F, v, w, Must, k  As Integer
Dim Exist(64), Work(), Finish, Critical As Boolean
Dim s, la As String
Dim u, u1, u2, u4, u7, uN, qq, qqq, ll, qL(0 To 76, 0 To 1), Gq(0 To 21), t, Over, Nq, Ad(0 To 21), q(0 To 64), HaveTo As Integer
Dim Force(0 To 21), Delay, GoUp, ex(0 To 64) As Boolean
Dim NextTo%(64)
Const Stick = "-"
Private Sub Command1_Click()
n = 1
Txt = Text1.Text
LoadStone
If Txt Like "???" Then
a = Mid(Txt, 1, 1)
b = Mid(Txt, 2, 1)
c = Mid(Txt, 3, 1)
If a < 5 And b < 5 And c < 5 And Not (a = 0 Or b = 0 Or c = 0) Then
If StoneLay(a, b, c) = 0 Then
PutStone a, b, c
Else
MsgBox ("The stone has already existed,try another place.")
End If
Else
MsgBox ("Please enter the number within 1 to 4.")
End If
Else
MsgBox ("Please enter the coordinate notation []-[]-[].")
End If
End Sub
Private Sub Command2_Click()
Txt = Text1.Text
If Txt Like "???" Then
a = Mid(Txt, 1, 1)
b = Mid(Txt, 2, 1)
c = Mid(Txt, 3, 1)
End If
StoneLay(a, b, c) = 0
End Sub
Private Sub PutStone(p1, p2, p3)
p = p1 + 4 * p2 + 16 * p3 - 20
StoneLay(p1, p2, p3) = ToGo
End Sub
Private Sub LoadStone()
For l = 0 To 76
    nL(l, -1) = 0
    nL(l, 1) = 0
 Next
For c = 1 To 4
  For b = 1 To 4
    For a = 1 To 4
     p = a + 4 * b + 16 * c - 20
     st = StoneLay(a, b, c)
     If StoneLay(a, b, c) = 0 Then
     Exist(p) = False
     Else
     Exist(p) = True
     For x = 1 To 7
     nL(cL(p, x), st) = nL(cL(p, x), st) + 1
     Next x
     End If
   Next a
  Next b
 Next c
End Sub
Private Sub Reflect()
For l = 0 To 76
    nL(l, -1) = 0
    nL(l, 1) = 0
 Next
Dim a, b, c As Integer
For p = 1 To 64
 cL(p, 1) = (p - 1) \ 4 + 1
 cL(p, 2) = 4 * ((p - 1) \ 16) + (p - 1) Mod 4 + 17
 cL(p, 3) = (p - 1) Mod 16 + 33
 Next
For a = 1 To 4
   For b = 1 To 4
     For c = 1 To 4
     p = a + 4 * b + 16 * c - 20
     If b = c Then
     cL(p, 4) = 47 + 2 * a
     ElseIf b + c = 5 Then
     cL(p, 4) = 48 + 2 * a
     End If
     If a = c Then
     cL(p, 5) = 55 + 2 * b
     ElseIf a + c = 5 Then
     cL(p, 5) = 56 + 2 * b
     End If
     If b = a Then
     cL(p, 6) = 63 + 2 * c
     ElseIf b + a = 5 Then
     cL(p, 6) = 64 + 2 * c
     End If
     If a = b And b = c Then
     cL(p, 7) = 73
     ElseIf a + b = 5 And b = c Then
     cL(p, 7) = 74
     ElseIf a = b And a + c = 5 Then
     cL(p, 7) = 75
     ElseIf a = c And a + b = 5 Then
     cL(p, 7) = 76
     End If
     Next
   Next
 Next
For p = 1 To 64
   For x = 1 To 7
   l = cL(p, x)
   If Not l = 0 Then
   nL(l, 1) = nL(l, 1) + 1
   sL(l, nL(l, 1)) = p
   End If
   Next
Next


  For p = 1 To 64
    NextTo(p) = p + 2 * (p Mod 2) - 1
  Next
  
End Sub

Private Sub Command3_Click()
Form1.Cls
Print ToGo
Turn = ToGo
LoadStone
If OnEarth(False, 20) = True Then
For x = 0 To 20
If Not Gq(x) = 0 Then
p = q(Gq(x)) - 1
p3 = p \ 16 + 1
p2 = (p Mod 16) \ 4 + 1
p1 = (p Mod 16) Mod 4 + 1
s = p1 & Stick & p2 & Stick & p3
Print s
End If
Next
Else
Print "No"
End If
End Sub

Private Sub Command4_Click()
StoneLay(1, 1, 4) = 1
StoneLay(1, 4, 4) = 1
StoneLay(4, 2, 2) = 1
StoneLay(3, 2, 2) = 1
StoneLay(2, 2, 2) = 1
StoneLay(4, 3, 1) = 1
StoneLay(2, 3, 1) = 1
StoneLay(1, 4, 1) = 1
StoneLay(4, 1, 4) = -1
StoneLay(1, 2, 4) = -1
StoneLay(1, 2, 2) = -1
StoneLay(1, 1, 1) = -1
StoneLay(4, 1, 1) = -1
StoneLay(4, 2, 1) = -1
StoneLay(4, 4, 1) = -1
ToGo = -1
Turn = -1
End Sub

Private Sub Form_Load()
ReDim StoneLay(4, 4, 4)
ToGo = 1
Reflect
End Sub

Private Sub Text2_Change()
ToGo = Val(Text2.Text)
End Sub
Private Function OnEarth(fc As Boolean, Endless As Integer)
k = 1
qq = 0
qqq = 0
Over = False
Delay = True
Force(1) = fc
Gq(1) = 0
For p = 1 To 64
ex(p) = Exist(p)
Next
qL(0, 1) = 0
qL(0, 0) = 0
For l = 1 To 76
qL(l, 1) = nL(l, Turn)
qL(l, 0) = nL(l, -Turn)
Next
uN = 0
For l = 1 To 76
  If qL(l, 1) = 2 And qL(l, 0) = 0 Then
    For u4 = 1 To 4
    If ex(sL(l, u4)) = False Then
    uN = uN + 1
    q(uN) = sL(l, u4)
    End If
    Next
  End If
Next
If uN = 0 Then
OnEarth = False
Exit Function
End If
Nq = uN
If Force(1) = False Then
  GoUp = False
Else
  GoUp = True
End If
Do
Working (Endless)
Loop Until k = 0
If Over = True Then
OnEarth = True
End If
End Function
Private Static Sub Working(Max As Integer)
If Gq(8) = 24 And Gq(9) = 5 And k = 9 Then Stop
If GoUp = True Then
  k = k + 1
  GoUp = False
  Force(k) = False
  If Force(k - 1) = True Then
  uN = 0
    For u = 1 To Nq
    If q(u) = HaveTo And ex(q(u)) = False And ex(q(NextTo(u))) = False Then
    Gq(k - 1) = u
    uN = uN + 1
    End If
    Next
    Select Case uN
    Case 0
    k = k - 1
    Exit Sub
    Case 2
    k = 0
    Over = True
    Exit Sub
    End Select
  End If
  If k = Max Then
  k = k - 1
  Exit Sub
  End If
  qq = Gq(k - 1)
  qqq = NextTo(qq)
  ex(q(qq)) = True
  ex(q(qqq)) = True
  uN = 0
    For u7 = 1 To 7
       l = cL(q(qqq), u7)
      If qL(l, 0) = 2 And qL(l, 1) = 0 And Not l = 0 Then
      uN = uN + 1
      For u4 = 1 To 4
      If ex(sL(l, u4)) = False Then
      HaveTo = sL(l, u4)
      End If
      Next
      End If
    Next
 Select Case uN
 Case 0
 Gq(k) = 0
 Case 1
  GoUp = True
  Force(k) = True
 Delay = True
 Case Else
 ex(q(qq)) = False
 ex(q(qqq)) = False
 k = k - 1
 Exit Sub
 End Select
 uN = 0
 For u7 = 1 To 7
  l = cL(q(qq), u7)
   u = cL(q(qqq), u7)
   qL(l, 1) = qL(l, 1) + 1
   qL(u, 0) = qL(u, 0) + 1
 Next
 For u7 = 1 To 7
   l = cL(q(qq), u7)
   If qL(l, 1) = 2 And qL(l, 0) = 0 And Not l = 0 Then
       For u4 = 1 To 4
       p = sL(l, u4)
       If ex(p) = False Then
         For u1 = 1 To Nq
           If p = q(u1) And Force(k) = False And ex(q(NextTo(u1))) = False Then
             k = 0
             Over = True
             Exit Sub
           End If
         Next
        uN = uN + 1
        q(Nq + uN) = p
      End If
      Next
   End If
 Next
Nq = Nq + uN
Ad(k) = uN
Exit Sub
End If
 
If Force(k) = False Then
     For u1 = 1 To Nq
       For u2 = u1 + 1 To Nq
       If q(u1) = q(u2) And ex(q(u1)) = False And ex(q(NextTo(u1))) = False Then
       k = 0
             Over = True
             Exit Sub
        End If
       Next
      Next
 For u = Gq(k) + 1 To Nq
   If ex(q(u)) = False And ex(q(NextTo(u))) = False Then
    Gq(k) = u
    GoUp = True
    Exit Sub
   End If
 Next
End If
  Nq = Nq - Ad(k)
k = k - 1
If k = 0 Then Exit Sub
u1 = Gq(k)
u2 = NextTo(u1)
ex(q(u1)) = False
ex(q(u2)) = False
For u7 = 1 To 7
  qL(cL(q(u1), u7), 1) = qL(cL(q(u1), u7), 1) - 1
  qL(cL(q(u2), u7), 0) = qL(cL(q(u2), u7), 0) - 1
  Next

End Sub
