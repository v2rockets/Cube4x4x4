Attribute VB_Name = "Match"
Option Base 1
Option Explicit
Public Gs0, Gs1, Gs2, Gsn2, Gsn1, Gt2, Gt3, Gtn1, Gtn2, Ge1, Ge2, Gen1, Gen2  As Single
Public Function PK(a() As Single) As Integer
Dim p, p1, p2, p3, n, Record, M As Integer
cN = 3: c1 = 8: c2 = 7: c3 = 3
Record = 0
ToGo = 1
For M = 1 To 2
Stone = 0
Erase StoneLay
Do
Gs0 = a(1, ToGo)
Gs1 = a(2, ToGo)
Gs2 = a(3, ToGo)
Gsn1 = a(4, ToGo)
Gsn2 = a(5, ToGo)
Gt2 = a(6, ToGo)
Gt3 = a(7, ToGo)
Gtn1 = a(8, ToGo)
Gtn2 = a(9, ToGo)
Ge1 = a(10, ToGo)
Ge2 = a(11, ToGo)
Gen1 = a(12, ToGo)
Gen2 = a(13, ToGo)
p = Compute
If Win(p) = False Then
  ValueX = Empty
  p = p - 1
  p3 = p \ 16 + 1
  p2 = (p Mod 16) \ 4 + 1
  p1 = (p Mod 16) Mod 4 + 1
  StoneLay(p1, p2, p3) = ToGo
  ToGo = -ToGo
  Stone = Stone + 1
Else
Record = Record + ToGo
Exit Do
End If
Loop Until CheckPath = True
ToGo = -1
Next
PK = Record / 2
End Function


