VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "read"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "write"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFileS As New FileSystemObject
Dim mFile As File
Dim mTxt As TextStream

Private Sub Command1_Click()
Set mFile = mFileS.GetFile("c:\t.dat")
Set mTxt = mFile.OpenAsTextStream(ForWriting)
Call mTxt.WriteLine(Text1.Text & "," & Text2.Text)

End Sub


Private Sub Command2_Click()
Dim s, s1, s2 As String
Dim mark1 As Integer
Set mTxt = mFile.OpenAsTextStream(ForReading)
s = mTxt.ReadLine
mark1 = InStr(1, s, ",")
s1 = Mid$(s, 1, mark1 - 1)
s2 = Mid$(s, mark1 + 1)
Print Val(s1) + Val(s2)

End Sub

Private Sub Command3_Click()
Form1.Cls
End Sub

Private Sub Form_Load()
Set mFile = mFileS.GetFile("c:\t.dat")
End Sub

Private Sub Form_Terminate()
mTxt.Close
End Sub
