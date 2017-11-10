VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cube4x4x4 v1.12(demo)"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Cube4x4x4new.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Choose depth:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim G(), VGo(), WGo(), StoneLay(), Turn, cL(64, 7), nL(0 To 76, -1 To 1), sL(0 To 76, 4) As Integer
Dim p, l, Depth, Stone, n, Suggest, ToGo, F, v, w, Must, k, Boundry(), tV() As Integer
Dim Exist(64), Work(), Finish, Critical, Legal As Boolean
Dim s, la As String
Dim u, u1, u2, u4, u7, uN, qq, qqq, ll, qL(0 To 76, 0 To 1), Gq(0 To 30), t, Over, Nq, Ad(0 To 30), q(0 To 64), HaveTo As Integer
Dim Force(0 To 30), Delay, GoUp, ex(0 To 64) As Boolean
Const Stick = "-"
Private Sub Choose()
Static x As Integer
If Critical = False Then
 If n = Depth Or n = 63 - Stone Then
    If OnEarth(False) = True Then
        n = n - 1
        Work(G(n), n) = True
        WGo(G(n), n) = (n + 2) * Turn
        Turn = -Turn
        UnDo (G(n))
        Exit Sub
    Else
        For x = 1 To 64
        If Exist(x) = False Then
        VGo(x, n) = Value(x)
        WGo(x, n) = 0
        End If
        Next
    End If
 Else
   If EvaT = False Then
   For x = 1 To 64
    If Work(x, n) = False And Exist(x) = False Then
      If tV(x, n) > Boundry Then
      GoPut x
      n = n + 1
      Exit Sub
      End If
    End If
  Next
  Else
  n = n - 1
        Work(G(n), n) = True
        WGo(G(n), n) = (n + 1) * Turn
        Turn = -Turn
        UnDo (G(n))
Else
  Critical = False
 If n = Depth Or n = 63 - Stone Then
    Exist(Must) = True
       fc = False
    For x = 1 To 7
      nL(cL(Must, x), Turn) = nL(cL(Must, x), Turn) + 1
    Next
    For x = 1 To 7
    fc = False
      If nL(cL(Must, x), -Turn) = 3 And nL(cL(Must, x), Turn) = 0 Then
      fc = True
         For u4 = 1 To 4
         p = sL(cL(Must, x), u4)
         If Exist(p) = False Then
         HaveTo = p
         End If
         Next
      Exit For
      End If
    Next
    Turn = -Turn
  If OnEarth(fc) = True Then
  UnDo (Must)
  n = n - 1
  WGo(G(n), n) = (n + 1) * Turn
  Else
  UnDo (Must)
  n = n - 1
  WGo(G(n), n) = 0
  VGo(G(n), n) = Value(Must)
  End If
  UnDo (G(n))
  Work(G(n), n) = True
  Exit Sub
 Else
    If EvaS(Must) = False Then
    n = n - 1
        Work(G(n), n) = True
        WGo(G(n), n) = (n + 1) * Turn
        Turn = -Turn
        UnDo (G(n))
    Else
    n = n + 1
 End If
  Exit Sub
End If
End If
 Compare
 n = n - 1
 Turn = -Turn
 UnDo (G(n))
 Work(G(n), n) = True
End Sub
Private Function EvaT() As Boolean
Static two, x, y, z, v, Vmax, Vmin As Integer
two = 0
Vmax = 0
Vmin = 100
For p = 1 To 64
If Exist(p, n) = False Then
 v = 0
 For x = 1 To 7
 l = cL(p, x)
  If nL(l, -Turn) = 0 Then
  Select Case nL(l, Turn)
  Case 2
  v = v + 8
  two = two + 1
  Case 1
  v = v + 3
  Case 0
  v = v + 1
  End Select
  ElseIf nL(l, Turn) = 0 Then
  Select Case nL(l, -Turn)
  Case 2
  v = v + 6
  Case 1
  v = v + 2
  End Select
  End If
 Next
  If v > Vmax Then v = Vmax
  If v < Vmin Then v = Vmin
  tV(p, n) = v
 End If
 Next
 If two > 1 Then
 EvaT = True
 Exit Function
 End If
 Boundry = (Vmax + Vmin) / 2
 EvaT = False
End Function
Private Function EvaS(m) As Boolean
Static two, x, y, z As Integer
two = 0
For x = 1 To 7
y = cL(m, x)
If (Not y = 0) And nL(y, -Turn) = 0 And nL(y, Turn) = 2 Then
two = two + 1
For z = 1 To 4
  If Exist(sL(y, z)) = False Then
  Must = sL(y, z)
  End If
Next
End If
Next
 Select Case two
 Case 0
 Critical = False
 Case 1
 Critical = True
 Case Else
 Critical = False
 EvaS = False
 Exit Function
 End Select
Exist(m) = True
G(n) = m
For x = 1 To 7
nL(cL(m, x), Turn) = nL(cL(m, x), Turn) + 1
Next
Turn = -Turn
 For x = 1 To 64
Work(x, n + 1) = False
Next
EvaS = True
End Function
Private Sub GoPut(m)
Critical = False
For x = 1 To 7
y = cL(m, x)
If (Not y = 0) And nL(y, -Turn) = 0 And nL(y, Turn) = 2 Then
For z = 1 To 4
  If Exist(sL(y, z)) = False Then
  Must = sL(y, z)
  Critical = True
  End If
Next
Exit For
End If
Next
Exist(m) = True
G(n) = m
For x = 1 To 7
nL(cL(m, x), Turn) = nL(cL(m, x), Turn) + 1
Next
Turn = -Turn
 For x = 1 To 64
Work(x, n + 1) = False
Next
End Sub

Private Function Win(m) As Boolean
Exist(m) = True
For x = 1 To 7
  nL(cL(m, x), Turn) = nL(cL(m, x), Turn) + 1
  Next
For l = 1 To 76
If nL(l, -1) = 4 Or nL(l, 1) = 4 Then
Win = True
UnDo (m)
Exit Function
End If
Next
UnDo (m)
Win = False
End Function


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
Dim p1, p2, p3 As Variant
Dim d%
d = Val(Text2.Text)
If d > 0 And d < 10 Then
Text1.Visible = False
Text2.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Label1.Caption = "Thinking...wait"
Label2.Visible = False
Label3.Visible = False
Label2.Caption = "Enter the notation:"
Depth = d
ReDim G(d)
ReDim WGo(64, d)
ReDim VGo(64, d)
ReDim Work(64, d)
ReDim Boundry(d)
Base
p = Suggest - 1
p3 = p \ 16 + 1
p2 = (p Mod 16) \ 4 + 1
p1 = (p Mod 16) Mod 4 + 1
s = p1 & Stick & p2 & Stick & p3
 If Finish = True Then
  If MsgBox("Just go " & s & " and win the game", vbOKCancel, " Computer") = vbOK Then
  Form_Load
  End If
  Else
   If MsgBox("I prefer to go " & s, vbOKCancel, " Computer") = vbOK Then
   PutStone p1, p2, p3
   Command2.Enabled = True
   Else
   Text1.Text = s
   Label2.Caption = "There is the hint:"
   End If
   Command3.Enabled = True
  End If
  Label1.Caption = la
  Text1.Visible = True
Text2.Visible = True
Command1.Enabled = True
Label2.Visible = True
Label3.Visible = True
Else
MsgBox ("Unavailable depth!")
End If
End Sub

Private Sub Command3_Click()
Form_Load
End Sub

Private Sub Command4_Click()
Command2.Enabled = True
End Sub

Private Sub Form_Load()
ToGo = 1
ReDim StoneLay(4, 4, 4)
Label2.Caption = "Enter the notation:"
la = "Red to play"
Label1.Caption = la
Text1.Text = Empty
Text2.Text = Empty
Command3.Enabled = False
Command2.Enabled = True
Reflect
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



Private Sub PutStone(p1, p2, p3)
p = p1 + 4 * p2 + 16 * p3 - 20
Turn = ToGo
If Win(p) = True Then
 If ToGo = 1 Then
 MsgBox ("Game over,red win!")
 Else
 MsgBox ("Game over,blue win!")
 End If
 Form_Load
ElseIf Stone = 63 Then
MsgBox ("Game end in a draw.")
  Else
  StoneLay(p1, p2, p3) = ToGo
  ToGo = -ToGo
  Command3.Enabled = True
  Stone = Stone + 1
  If ToGo = 1 Then
 la = "Red to play"
 Label1.Caption = la
  Else
   la = "Blue to play"
   Label1.Caption = la
   End If
   End If
Command2.Enabled = True
End Sub



Private Sub Base()
Dim wait(64), m, z As Integer
Static x%
Finish = False
Critical = False
n = 1
 LoadStone
Turn = ToGo
For x = 1 To 76
  If nL(x, ToGo) = 0 And nL(x, -ToGo) = 3 Then
  Critical = True
  For z = 1 To 4
    If Exist(sL(x, z)) = False Then
    Suggest = sL(x, z)
    Exit For
    End If
  Next
  End If
Next
For x = 1 To 64
  If Exist(x) = False Then
    If Win(x) = True Then
     Finish = True
     Suggest = x
     Exit Sub
  End If
 End If
Next
If Critical = True Then
  Exit Sub
Else
   For x = 1 To 64
     If Exist(x) = False Then
     WGo(x, n) = 0
     If (Stone + n) = 63 Then
     VGo(x, n) = Value(x)
     Else
       If EvaS(x) = True Then
       n = n + 1
          Do
          Choose
          Loop Until n = 1
       Else
       Suggest = x
       Exit Sub
       End If
     End If
     End If
     Next
End If
  If ToGo = 1 Then
  w = -64
  For p = 1 To 64
    If Exist(p) = False And WGo(p, 1) > w Then
    w = WGo(p, 1)
    End If
  Next
  If w = 0 Then
    v = -1024
    For p = 1 To 64
    If Exist(p) = False And WGo(p, 1) = 0 And VGo(p, 1) > v Then
    v = VGo(p, 1)
    End If
    Next
    
  End If
Else
  w = 64
  For p = 1 To 64
    If Exist(p) = False And WGo(p, 1) < w Then
    w = WGo(p, 1)
    End If
  Next
  If w = 0 Then
    v = 1024
    For p = 1 To 64
    If Exist(p) = False And WGo(p, 1) = 0 And VGo(p, 1) < v Then
    v = VGo(p, 1)
    End If
    Next
    
  End If
End If
 If w = 0 Then
 m = 0
 For x = 1 To 64
 If Exist(x) = False And VGo(x, 1) = v And WGo(x, 1) = 0 Then
 m = m + 1
 wait(m) = x
 End If
 Next x
Randomize
m = Rnd() * m + 0.5
 Suggest = wait(m)
 Else
 For x = 1 To 64
 If Exist(x) = False And WGo(x, 1) = w Then
 Suggest = x
 Exit Sub
 End If
 Next
 End If
 End Sub
Static Sub Compare()
If Turn = 1 Then
w = -63
  For p = 1 To 64
    If Exist(p) = False And WGo(p, n) > w Then
    w = WGo(p, n)
    End If
  Next
  If w = 0 Then
    v = -1023
    For p = 1 To 64
    If Exist(p) = False And WGo(p, n) = 0 And VGo(p, n) > v Then
    v = VGo(p, n)
    End If
    Next
    VGo(G(n - 1), n - 1) = v
    
  Else
  WGo(G(n - 1), n - 1) = w
  End If
Else
  w = 63
  For p = 1 To 64
    If Exist(p) = False And WGo(p, n) < w Then
    w = WGo(p, n)
    End If
  Next
  If w = 0 Then
    v = 1023
    For p = 1 To 64
    If Exist(p) = False And WGo(p, n) = 0 And VGo(p, n) < v Then
    v = VGo(p, n)
    End If
    Next
    VGo(G(n - 1), n - 1) = v
    WGo(G(n - 1), n - 1) = 0
  Else
  WGo(G(n - 1), n - 1) = w
  End If
End If
End Sub

Private Sub Text1_GotFocus()
Label2.Caption = "Enter the notation:"
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
End Sub

Private Sub UnDo(m As Integer)
Static x%
Exist(m) = False
For x = 1 To 7
nL(cL(m, x), Turn) = nL(cL(m, x), Turn) - 1
Next
End Sub

Private Function OnEarth(fc)
OnEarth = False
End Function

Private Function Value(m)
v = 0
Exist(m) = True
For x = 1 To 7
  nL(cL(m, x), Turn) = nL(cL(m, x), Turn) + 1
  Next
For l = 1 To 76
If nL(l, -1) = 0 Then
Select Case nL(l, 1)
Case 3
v = v + 12
Case 2
v = v + 4
Case 1
v = v + 1
End Select
ElseIf nL(l, 1) = 0 Then
Select Case nL(l, -1)
Case 3
v = v - 12
Case 2
v = v - 4
Case 1
v = v - 1
End Select
End If
Next
UnDo (m)
Value = v
End Function

