VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GAOS"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6885
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Birth 
      Caption         =   "new"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox Check 
      Caption         =   "autowork"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Timer Auto 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4680
      Top             =   600
   End
   Begin VB.CommandButton Combine 
      Caption         =   "combine"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton ReadInfo 
      Caption         =   "read"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Part2 
      Caption         =   "do part2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Part1 
      Caption         =   "do part1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "do all"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Dim GFunctor(13, 65) As Single
Dim mFileS As New FileSystemObject
Dim mFile As File
Dim mTxt As TextStream
Dim Info As String
Dim Times, Amount, Part, Position As Variant

Private Sub Birth_Click()
Dim Gtemp(13) As Integer
Dim n%, f%, t%
Gtemp(1) = 3 / 3
Gtemp(2) = 7 / 3
Gtemp(3) = 10 / 3
Gtemp(4) = 4 / 3
Gtemp(5) = 14 / 3
Gtemp(6) = 10 / 3
Gtemp(7) = 20 / 3
Gtemp(8) = 4 / 3
Gtemp(9) = 14 / 3
Gtemp(10) = 7 / 3
Gtemp(11) = 10 / 3
Gtemp(12) = 4 / 3
Gtemp(13) = 14 / 3
For f = 1 To 64
  For n = 1 To 13
  GFunctor(n, f) = Gtemp(n) * 2 ^ (2 * Rnd - 1)
  Next
Next
Times = 0
Amount = "64*13"
Position = 0
Part = "1/1"
mFileS.CreateTextFile ("c:\GA.cdb")
Set mFile = mFileS.GetFile("c:\GA.cdb")
SaveData
End Sub

Private Sub Check_Click()
Auto.Enabled = True
End Sub

Private Sub Combine_Click()
Dim n%, f%
Dim Times0%
Dim Ftemp(13, 32) As Single
Set mFile = mFileS.GetFile("c:\GA.cdb")
Read
If Not (Part = "1/1" And Position = 0) Then MsgBox ("error0"): End
Times0 = Times
Set mFile = mFileS.GetFile(("c:\GAtemp1.cdb"))
Read
If Not (Times = Times0 And Part = "1/2" And Position = 16 And Amount = "64*13") Then MsgBox ("error1"): End
For f = 1 To 32
For n = 1 To 13
Ftemp(n, f) = GFunctor(n, f)
Next
Next
Set mFile = mFileS.GetFile(("c:\GAtemp2.cdb"))
Read
If Not (Times = Times0 And Part = "2/2" And Position = 32 And Amount = "64*13") Then MsgBox ("error2"): End
For f = 1 To 32
For n = 1 To 13
GFunctor(n, f) = Ftemp(n, f)
Next
Next
MoveG
Times = Times0 + 1
Part = "1/1"
Position = 0
Set mFile = mFileS.GetFile("c:\GA.cdb")
SaveData
Kill "c:\GAtemp1.cdb"
Kill "c:\GAtemp2.cdb"
MsgBox ("success")
End Sub

Private Sub Command1_Click()
If Not (Part = "1/1" And Position = 0) Then MsgBox ("error"): Exit Sub
SaveData
End Sub
Private Sub Command2_Click()
Dim n%, f%, t%
Static Temp(13, -1 To 1) As Single
Form1.Hide
For f = 1 To 32
For n = 1 To 13
Temp(n, 1) = GFunctor(n, 2 * f - 1)
Temp(n, -1) = GFunctor(n, 2 * f)
Next
t = PK(Temp())
If t = 0 Then
For n = 1 To 3
Exchange GFunctor(), 2 * f - 1, GFunctor(), 2 * f
Next
Else
For n = 1 To 13
If t = 1 Then
GFunctor(n, 2 * f) = GFunctor(n, 2 * f - 1)
Else
GFunctor(n, 2 * f - 1) = GFunctor(n, 2 * f)
End If
Next
End If
Vary GFunctor(), 2 * f - 1
Vary GFunctor(), 2 * f
Next
MoveG
Times = Times + 1
Position = 0
Form1.Show
Auto.Enabled = True
End Sub

Private Sub Part2_Click()
Dim n%, f%, t%
Static Temp(13, -1 To 1) As Single
mFileS.CreateTextFile ("c:\GAtemp2.cdb")
Command2.Enabled = False
Command1.Enabled = False
For f = 17 To 32
For n = 1 To 13
Temp(n, 1) = GFunctor(n, 2 * f - 1)
Temp(n, -1) = GFunctor(n, 2 * f)
Next
t = PK(Temp())
If t = 0 Then
For n = 1 To 3
Exchange GFunctor(), 2 * f - 1, GFunctor(), 2 * f
Next
Else
For n = 1 To 13
If t = 1 Then
GFunctor(n, 2 * f) = GFunctor(n, 2 * f - 1)
Else
GFunctor(n, 2 * f - 1) = GFunctor(n, 2 * f)
End If
Next
End If
Vary GFunctor(), 2 * f - 1
Vary GFunctor(), 2 * f
Next
If MsgBox("Ready to save", vbYesNo) = vbYes Then
Part = "2/2"
Position = 32
Set mFile = mFileS.GetFile(("c:\GAtemp2.cdb"))
SaveData
Part2.Enabled = False
Else
End
End If
End
End Sub

Private Sub ReadInfo_Click()
Dim k%
If Not (Part = "1/1" And Position = 0 And Amount = "64*13") Then MsgBox ("error"): End
Form1.Cls
Print Info
For k = 1 To 13
Print "C" & k & ":" & CalAveSd(GFunctor(), k)
Next
Command2.Enabled = True
Part1.Enabled = True
Part2.Enabled = True
Check.Enabled = True
End Sub


Private Sub Part1_Click()
Dim n%, f%, t%
Static Temp(13, -1 To 1) As Single
mFileS.CreateTextFile ("c:\GAtemp1.cdb")
Command2.Enabled = False
Command1.Enabled = False
For f = 1 To 16
For n = 1 To 13
Temp(n, 1) = GFunctor(n, 2 * f - 1)
Temp(n, -1) = GFunctor(n, 2 * f)
Next
t = PK(Temp())
If t = 0 Then
For n = 1 To 3
Exchange GFunctor(), 2 * f - 1, GFunctor(), 2 * f
Next
Else
For n = 1 To 13
If t = 1 Then
GFunctor(n, 2 * f) = GFunctor(n, 2 * f - 1)
Else
GFunctor(n, 2 * f - 1) = GFunctor(n, 2 * f)
End If
Next
End If
Vary GFunctor(), 2 * f - 1
Vary GFunctor(), 2 * f
Next
If MsgBox("Ready to save", vbYesNo) = vbYes Then
Part = "1/2"
Position = 16
Set mFile = mFileS.GetFile(("c:\GAtemp1.cdb"))
SaveData
Part1.Enabled = False
Else
End
End If
End
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Reflect
Set mFile = mFileS.GetFile("c:\GA.cdb")
Read
End Sub

Private Sub Vary(a() As Single, f As Integer)
Dim n%
Randomize
For n = 1 To 13
a(n, f) = a(n, f) * 0.99 ^ (2 * Rnd() - 1)
Next
End Sub

Private Sub Read()
Dim n%, f%
Set mTxt = mFile.OpenAsTextStream(ForReading)
Dim s, s1, s2 As String
Dim mark1, mark2 As Integer
Set mTxt = mFile.OpenAsTextStream(ForReading)
Info = mTxt.ReadLine
If Not Mid$(Info, 1, 9) = "<StdInfo>" Then End
GetInfo
For f = 1 To 64
s = mTxt.ReadLine
mark2 = 0
  For n = 1 To 13
  mark1 = mark2
  mark2 = InStr(mark2 + 1, s, " ")
  GFunctor(n, f) = Val(Mid(s, mark1 + 1, mark2 - mark1 - 1))
  Next
Next
mTxt.Close
End Sub

Private Sub GetInfo()
Dim mark1, mark2 As Integer
mark1 = InStr(1, Info, "Amount")
mark1 = InStr(mark1, Info, ":") + 1
mark2 = InStr(mark1, Info, " ")
Amount = Mid(Info, mark1, mark2 - mark1)

mark1 = InStr(1, Info, "Times")
mark1 = InStr(mark1, Info, ":") + 1
mark2 = InStr(mark1, Info, " ")
Times = Mid(Info, mark1, mark2 - mark1)

mark1 = InStr(1, Info, "Part")
mark1 = InStr(mark1, Info, ":") + 1
mark2 = InStr(mark1, Info, " ")
Part = Mid(Info, mark1, mark2 - mark1)

mark1 = InStr(1, Info, "Position")
mark1 = InStr(mark1, Info, ":") + 1
mark2 = InStr(mark1, Info, " ")
Position = Mid(Info, mark1, mark2 - mark1)
End Sub

Private Sub Exchange(a1() As Single, f1 As Integer, a2() As Single, f2 As Integer)
Dim Gtemp As Single
Dim R%, way%
Randomize
R = Round(13 * Rnd + 0.5)
way = Round(Rnd)
If way = 0 Then
Gtemp = a1(R, f1)
a1(R, f1) = a2(R, f2)
a2(R, f2) = Gtemp
Else
Gtemp = (a1(R, f1) + a2(R, f2)) / 2
a1(R, f1) = Gtemp
a2(R, f2) = Gtemp
End If
End Sub

Private Sub MoveG()
Dim n%, f%
For f = 63 To 1 Step -2
  For n = 1 To 13
 GFunctor(n, f + 2) = GFunctor(n, f)
  Next
Next
For n = 1 To 13
 GFunctor(n, 1) = GFunctor(n, 65)
 Next
End Sub

Private Sub SaveData()
Dim n%, f%
Set mTxt = mFile.OpenAsTextStream(ForWriting)
Call mTxt.WriteLine("<StdInfo>Amount:" & Amount & " Times:" & Times & " Part:" & Part & " Position:" & Position & " |")
For f = 1 To 64
Call mTxt.WriteLine(GFunctor(1, f) & " " & GFunctor(2, f) & " " & GFunctor(3, f) & " " & GFunctor(4, f) & " " & GFunctor(5, f) & " " & GFunctor(6, f) & " " & GFunctor(7, f) & " " & GFunctor(8, f) & " " & GFunctor(9, f) & " " & GFunctor(10, f) & " " & GFunctor(11, f) & " " & GFunctor(12, f) & " " & GFunctor(13, f) & " |")
Next
mTxt.Close
End Sub

Private Sub Auto_Timer()
If Check.Value = 0 Then
Check.Enabled = False
Auto.Enabled = False
Part1.Enabled = False
Part2.Enabled = False
Command1.Enabled = True
Exit Sub
Else
Command1.Enabled = False
Command2.Enabled = False
Command2_Click
End If
End Sub

Private Function CalAveSd(GF() As Single, n As Integer) As String
Dim a, ave, b, sd, logGF(64) As Single
Dim f%
a = 0
For f = 1 To 64
logGF(f) = Log(GF(n, f))
a = a + logGF(f)
Next
a = a / 64
ave = Exp(a)
b = 0
For f = 1 To 64
b = b + (logGF(f) - a) ^ 2
Next
b = b / 64
sd = Sqr(b)
CalAveSd = ave & "," & sd
End Function

