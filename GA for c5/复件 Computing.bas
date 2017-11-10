Attribute VB_Name = "Computing"
Option Base 1
Option Explicit
Dim u, u1, u2, u4, u7, uN, qq, qqq, ll, qL(0 To 76, 0 To 1), Gq(0 To 30), t, Over, Nq, Ad(0 To 30), q(0 To 64), HaveTo As Integer
Dim Work(64, 10), Finish, Critical(10), News, AddDepth(10) As Boolean
Dim G(10), VGo(10) As Integer
Dim p, l, Depth, n, Suggest, f, v, w, Must, k, Boundry(10), tV(64, 10), WP, Near(60) As Integer
Dim Force(0 To 30), Delay, GoUp, ex(0 To 64) As Boolean
Dim Ghost(64), st, nd As Integer
Public ValueX As Integer
Private Sub Choose()
Static x, Temp, vv As Integer
Static fc As Boolean
If News = True Then
News = False
If Critical(n - 1) = False Then
    If OnEarth(Critical(n), c3 + Depth - n) = True Then
        VGo(n) = 1000 - n - WP
         n = n - 1
       Work(G(n), n) = True
        Turn = -Turn
        UnDo (G(n))
     If -VGo(n + 1) > VGo(n) Then VGo(n) = -VGo(n + 1)
     Exit Sub
      End If
   End If
  If n + Stone = 63 Or n = Depth Then
      If Critical(n) = True Then
      VGo(n) = ValueT + ValueS(Must)
      Else
      VGo(n) = Value + ValueT
      End If
       n = n - 1
       Work(G(n), n) = True
        Turn = -Turn
        UnDo (G(n))
     If -VGo(n + 1) > VGo(n) Then VGo(n) = -VGo(n + 1)
     Exit Sub
  Else
   If Critical(n) = True Then
   VGo(n) = VGo(n - 2) - 1
   GoPut Must
   If n < cN Then
   AddDepth(n) = True
   Depth = Depth + 2
   End If
   n = n + 1
   Exit Sub
   End If
    VGo(n) = -1000
   EvaT
  End If
End If

If Critical(n) = False Then
  If VGo(n - 1) > -VGo(n) Or (VGo(n - 1) = -VGo(n) And Not n = 2) Then
  If AddDepth(n) = True Then
  AddDepth(n) = False
  Depth = Depth - 2
  End If
   n = n - 1
    Work(G(n), n) = True
    Turn = -Turn
    UnDo (G(n))
    Exit Sub
  End If
  v = 0
   For x = 1 To 64
   If Exist(x) = False And Work(x, n) = False Then
   vv = tV(x, n)
      If vv > v Then
      v = vv
      Temp = x
      End If
   End If
   Next
  If Not v = 0 Then
  
  GoPut Temp
  n = n + 1
  Exit Sub
  End If
End If
If AddDepth(n) = True Then
AddDepth(n) = False
Depth = Depth - 2
End If
n = n - 1
    Work(G(n), n) = True
    Turn = -Turn
    UnDo (G(n))
 If -VGo(n + 1) > VGo(n) Then VGo(n) = -VGo(n + 1)
End Sub
Private Sub Virsual()
Static x%
st = 0
nd = 0
For p = 1 To 64
  If Exist(p) = False Then
  v = 0
  For x = 1 To 7
     l = cL(p, x)
     If Not l = 0 Then
      If nL(l, -Turn) = 0 Then
      Select Case nL(l, Turn)
      Case 0
        v = v + Gg0
        Case 1
         v = v + Gg1
        Case 2
         v = v + Gg2
      End Select
     ElseIf nL(l, Turn) = 0 Then
        Select Case nL(l, -Turn)
        Case 1
         v = v + Ggn1
        Case 2
         v = v + Ggn2
      End Select
     End If
    End If
 Next
 Ghost(p) = v
 If v > st Then
 st = v
 nd = st
 ElseIf v > nd Then
 nd = v
 End If
 End If
Next
End Sub
Private Sub EvaT()
Static x, v, vmax, vmin As Integer
Virsual
vmax = 0
vmin = 100
For p = 1 To 64
   If Exist(p) = False Then
   v = 0
    For x = 1 To 7
     l = cL(p, x)
     If Not l = 0 Then
      If nL(l, -Turn) = 0 Then
      Select Case nL(l, Turn)
      Case 0
        v = v + 3
        Case 1
         v = v + Ge12
        Case 2
         v = v + Ge23
      End Select
     ElseIf nL(l, Turn) = 0 Then
        Select Case nL(l, -Turn)
        Case 1
        v = v + Gen1
        Case 2
        v = v + Gen2
      End Select
     End If
    End If
    Next
    If Ghost(p) = st Then
    v = v - nd
    Else
    v = v - st
    End If
  tV(p, n) = v
 End If
 Next
End Sub

Private Sub GoPut(m)
Static x%, y%, z%, p%
p = m
Critical(n + 1) = False
Exist(p) = True
For x = 1 To 7
y = cL(p, x)
If (Not y = 0) And nL(y, -Turn) = 0 And nL(y, Turn) = 2 Then
For z = 1 To 4
  If Exist(sL(y, z)) = False Then
  Must = sL(y, z)
  Critical(n + 1) = True
  End If
Next
Exit For
End If
Next
G(n) = p
For x = 1 To 7
nL(cL(p, x), Turn) = nL(cL(p, x), Turn) + 1
Next
Turn = -Turn
News = True
 For x = 1 To 64
Work(x, n + 1) = False
Next
End Sub
Private Sub UnDo(m As Integer)
Static x%
Exist(m) = False
For x = 1 To 7
nL(cL(m, x), Turn) = nL(cL(m, x), Turn) - 1
Next
End Sub
Private Function OnEarth(fc, Endless) As Boolean
Dim Km As Integer
Km = Endless
GoUp = False
If fc = True Then
uN = 0
For u7 = 1 To 7
l = cL(Must, u7)
If nL(l, Turn) = 2 And nL(l, -Turn) = 0 And Not l = 0 Then uN = uN + 1
Next
Select Case uN
Case 0
OnEarth = False
Exit Function
Case 1
  GoUp = True
  HaveTo = Must
Case Else
OnEarth = True
Exit Function
End Select
End If
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
Do
Working Km
Loop Until k = 0
If Over = True Then
OnEarth = True
End If
End Function
Private Static Sub Working(Km As Integer)
If GoUp = True Then
  k = k + 1
  GoUp = False
  Force(k) = False
  If Force(k - 1) = True Then
  uN = 0
    For u = 1 To Nq
    If q(u) = HaveTo And ex(q(u)) = False And ex(q(Near(u))) = False Then
    Gq(k - 1) = u
    uN = uN + 1
    End If
    Next
    Select Case uN
    Case 0
    k = k - 1
    Exit Sub
    Case 2
    WP = k - 1
    k = 0
    Over = True
    Exit Sub
    End Select
  End If
  If k > Km Then
  k = k - 1
  Exit Sub
  End If
  qq = Gq(k - 1)
  qqq = Near(qq)
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
           If p = q(u1) And Force(k) = False Then
           WP = k
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
   If Delay = True Then
    Delay = False
     For u1 = 1 To Nq
       For u2 = u1 + 1 To Nq
       If q(u1) = q(u2) And ex(q(u1)) = False And ex(q(Near(u1))) = False And ex(q(Near(u2))) = False Then
             Over = True
             WP = k
            k = 0
             Exit Sub
        End If
       Next
      Next
   End If
 If Force(k - 1) = False And Gq(k) < Gq(k - 1) Then
 For u = Gq(k) + 1 To Gq(k - 1) - 1
    If ex(q(u)) = False And ex(q(Near(u))) = False Then
    For u7 = 1 To 7
       l = cL(q(Near(u)), u7)
      If qL(l, 0) = 2 And qL(l, 1) = 0 And Not l = 0 Then uN = uN + 1
    Next
    If uN = 1 Then
    Gq(k) = u
    GoUp = True
    Exit Sub
    End If
    End If
 Next
 Gq(k) = Gq(k - 1)
 End If
 For u = Gq(k) + 1 To Nq
   If ex(q(u)) = False And ex(q(Near(u))) = False Then
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
u2 = Near(u1)
ex(q(u1)) = False
ex(q(u2)) = False
For u7 = 1 To 7
  qL(cL(q(u1), u7), 1) = qL(cL(q(u1), u7), 1) - 1
  qL(cL(q(u2), u7), 0) = qL(cL(q(u2), u7), 0) - 1
  Next

End Sub

Private Function ValueT()
Static x, v As Integer
v = 0
For l = 1 To 76
If nL(l, -Turn) = 0 Then
Select Case nL(l, Turn)
      Case 1
        v = v + 3
        Case 2
        v = v + Gv2
End Select
ElseIf nL(l, Turn) = 0 Then
Select Case nL(l, -Turn)
    Case 1
     v = v - 3
    Case 2
     v = v - Gv2
End Select
End If
Next
ValueT = v
End Function
Private Function ValueS(m)
Static x, v, vmax As Integer
Virsual
If Ghost(m) = st Then
v = -nd
Else
v = -st
End If
For x = 1 To 7
     l = cL(m, x)
     If Not l = 0 Then
      If nL(l, -Turn) = 0 Then
      Select Case nL(l, Turn)
      Case 0
        v = v + 3
        Case 1
         v = v + Gv12
        Case 2
         v = v + Gv23
      End Select
     ElseIf nL(l, Turn) = 0 Then
        Select Case nL(l, -Turn)
        Case 1
        v = v + 3
        Case 2
        v = v + Gv23
      End Select
     End If
    End If
Next
ValueS = v
End Function

Private Function Value()
Static x, v, vmax, m As Integer
Virsual
vmax = -256
For m = 1 To 64
If Exist(m) = False Then
  v = 0
  For x = 1 To 7
     l = cL(m, x)
     If Not l = 0 Then
      If nL(l, -Turn) = 0 Then
      Select Case nL(l, Turn)
      Case 0
        v = v + 3
        Case 1
         v = v + Gv12
        Case 2
         v = v + Gv23
      End Select
     ElseIf nL(l, Turn) = 0 Then
        Select Case nL(l, -Turn)
        Case 1
         v = v + 3
        Case 2
         v = v + Gv2
      End Select
     End If
    End If
 Next
 If Ghost(m) = st Then
 v = v - nd
 Else
 v = v - st
 End If
 If v > vmax Then vmax = v
End If
Next
Value = vmax
End Function
Public Function Compute()
Dim d%, m%
Dim a, b, C As Integer
d = cN
Depth = d
LoadStone
For p = 1 To 60
Near(p) = p + 2 * (p Mod 2) - 1
Next
If Stone = 0 Then
Randomize
Do
m = 64 * Rnd() + 0.5
p = m - 1
C = p \ 16 + 1
b = (p Mod 16) \ 4 + 1
a = (p Mod 16) Mod 4 + 1
Loop Until Not (((a = 1 Or a = 4) And (b = 1 Or b = 4) And (C = 1 Or C = 4)) Or ((a = 2 Or a = 3) And (b = 2 Or b = 3) And (C = 2 Or C = 3)))
Compute = m
Else
Base
Compute = Suggest
End If
End Function

Private Function Reduce(Impo, deep) As Boolean
Dim Im As Boolean
HaveTo = Must
Im = Impo
WP = deep
If OnEarth(Im, WP) = True Then
Reduce = True
Do
Suggest = q(Gq(1))
WP = WP - 1
If WP < 1 Then Exit Do
Loop While OnEarth0(Im, WP) = True
WP = WP + 1
Else
Reduce = False
End If
End Function
Private Sub Base()
Dim wait(64), mm, z, Choice, VChoice(64) As Integer
Static x, vv As Integer
Finish = False
Critical(1) = False
n = 1
Turn = ToGo
If Stone = 63 Then
  For x = 1 To 64
  If Exist(x) = True Then Suggest = x
  Next
  Exit Sub
End If
For x = 1 To 64
Work(x, 1) = False
VChoice(x) = -1000
Next
For x = 1 To 76
  If nL(x, ToGo) = 0 And nL(x, -ToGo) = 3 Then
  Critical(1) = True
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
If Critical(1) = True Then
  Exit Sub
Else
 If Reduce(False, c1) = True Then
   For p = 1 To 64
    uN = 0
    If Exist(p) = False Then
     For z = 1 To 7
     l = cL(p, z)
     If Not l = 0 Then
        If nL(l, -ToGo) = 0 And nL(l, ToGo) = 2 Then uN = uN + 1
     End If
     Next
      If uN > 1 Then
      Suggest = p
      End If
    End If
  Next
  ValueX = 1000 - WP * 2
Exit Sub
 End If
        VGo(n) = -1000
        EvaT
  Do
    vv = 0
     For p = 1 To 64
      If Exist(p) = False And Work(p, 1) = False Then
      z = tV(p, 1)
        If z > vv Then
        vv = z
        x = p
        End If
      End If
     Next
     If Not vv = 0 Then
         GoPut x
         If Reduce(Critical(2), c2) = False Then
         n = n + 1
         News = False
         VGo(n) = -1000
           If Critical(n) = True Then
           GoPut Must
            AddDepth(n) = True
            Depth = Depth + 2
            n = n + 1
           Else
           EvaT
           End If
          Do
          Choose
          Loop Until n = 1
          VChoice(x) = -VGo(2)
         Else
           VChoice(x) = 2 * WP - 1000
          Turn = -Turn
         UnDo (x)
         Work(x, 1) = True
          If VChoice(x) > VGo(1) Then VGo(1) = VChoice(x)
         End If
    End If
  Loop Until vv = 0
End If
mm = 0
For x = 1 To 64
 If Exist(x) = False And VChoice(x) = VGo(1) Then
 mm = mm + 1
 wait(mm) = x
 End If
Next
  Randomize
  mm = Rnd() * mm + 0.5
  Suggest = wait(mm)
  ValueX = VGo(1)
End Sub
Private Function OnEarth0(fc, Endless) As Boolean
Dim Km As Integer
Km = Endless
GoUp = False
If fc = True Then
uN = 0
For u7 = 1 To 7
l = cL(Must, u7)
If nL(l, Turn) = 2 And nL(l, -Turn) = 0 And Not l = 0 Then uN = uN + 1
Next
Select Case uN
Case 0
OnEarth0 = False
Exit Function
Case 1
  GoUp = True
  HaveTo = Must
Case Else
OnEarth0 = True
WP = 1
Exit Function
End Select
End If
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
OnEarth0 = False
Exit Function
End If
Nq = uN
Do
Working0 Km
Loop Until k = 0
If Over = True Then
OnEarth0 = True
End If
End Function
Private Static Sub Working0(Km As Integer)
If GoUp = True Then
  k = k + 1
  GoUp = False
  Force(k) = False
  If Force(k - 1) = True Then
  uN = 0
    For u = 1 To Nq
    If q(u) = HaveTo And ex(q(u)) = False And ex(q(Near(u))) = False Then
    Gq(k - 1) = u
    uN = uN + 1
    End If
    Next
    Select Case uN
    Case 0
    k = k - 1
    Exit Sub
    Case 2
    WP = k - 1
    k = 0
    Over = True
    Exit Sub
    End Select
  End If
  If k > Km Then
  k = k - 1
  Exit Sub
  End If
  qq = Gq(k - 1)
  qqq = Near(qq)
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
           If p = q(u1) And Force(k) = False Then
           WP = k
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
   If Delay = True Then
    Delay = False
     For u1 = 1 To Nq
       For u2 = u1 + 1 To Nq
       If q(u1) = q(u2) And ex(q(u1)) = False And ex(q(Near(u1))) = False And ex(q(Near(u2))) = False Then
             Over = True
             WP = k
              k = 0
             Exit Sub
        End If
       Next
      Next
   End If
 For u = Gq(k) + 1 To Nq
   If ex(q(u)) = False And ex(q(Near(u))) = False Then
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
u2 = Near(u1)
ex(q(u1)) = False
ex(q(u2)) = False
For u7 = 1 To 7
  qL(cL(q(u1), u7), 1) = qL(cL(q(u1), u7), 1) - 1
  qL(cL(q(u2), u7), 0) = qL(cL(q(u2), u7), 0) - 1
  Next

End Sub

