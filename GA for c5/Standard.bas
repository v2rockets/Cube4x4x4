Attribute VB_Name = "Stantard"
Dim p, l, x As Integer
Public StoneLay(4, 4, 4), cL(64, 7), nL(0 To 76, -1 To 1), sL(0 To 76, 4), ToGo, Stone, Turn As Integer
Public Exist(64) As Boolean
Public c1, c2, c3, cN, winL As Integer
Dim a%, b%, C%
Public Function Win(m) As Boolean
Exist(m) = True
For x = 1 To 7
  nL(cL(m, x), Turn) = nL(cL(m, x), Turn) + 1
  Next
For l = 1 To 76
If nL(l, -1) = 4 Or nL(l, 1) = 4 Then
Win = True
winL = l
UnDo (m)
Exit Function
End If
Next
UnDo (m)
Win = False
End Function
Public Sub LoadStone()
For l = 0 To 76
    nL(l, -1) = 0
    nL(l, 1) = 0
 Next
For C = 1 To 4
  For b = 1 To 4
    For a = 1 To 4
     p = a + 4 * b + 16 * C - 20
     st = StoneLay(a, b, C)
     If StoneLay(a, b, C) = 0 Then
     Exist(p) = False
     Else
     Exist(p) = True
     For x = 1 To 7
     nL(cL(p, x), st) = nL(cL(p, x), st) + 1
     Next x
     End If
   Next a
  Next b
 Next C
End Sub

Public Sub Reflect()
For l = 0 To 76
    nL(l, -1) = 0
    nL(l, 1) = 0
 Next
Dim a, b, C As Integer
For p = 1 To 64
 cL(p, 1) = (p - 1) \ 4 + 1
 cL(p, 2) = 4 * ((p - 1) \ 16) + (p - 1) Mod 4 + 17
 cL(p, 3) = (p - 1) Mod 16 + 33
 Next
For a = 1 To 4
   For b = 1 To 4
     For C = 1 To 4
     p = a + 4 * b + 16 * C - 20
     If b = C Then
     cL(p, 4) = 47 + 2 * a
     ElseIf b + C = 5 Then
     cL(p, 4) = 48 + 2 * a
     End If
     If a = C Then
     cL(p, 5) = 55 + 2 * b
     ElseIf a + C = 5 Then
     cL(p, 5) = 56 + 2 * b
     End If
     If b = a Then
     cL(p, 6) = 63 + 2 * C
     ElseIf b + a = 5 Then
     cL(p, 6) = 64 + 2 * C
     End If
     If a = b And b = C Then
     cL(p, 7) = 73
     ElseIf a + b = 5 And b = C Then
     cL(p, 7) = 74
     ElseIf a = b And a + C = 5 Then
     cL(p, 7) = 75
     ElseIf a = C And a + b = 5 Then
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

Public Function CheckPath() As Boolean
For l = 1 To 76
If nL(l, 1) = 0 Or nL(l, -1) = 0 Then Exit Function
Next
CheckPath = True
End Function
