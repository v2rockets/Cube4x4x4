VERSION 5.00
Begin VB.Form Visual 
   AutoRedraw      =   -1  'True
   Caption         =   "Cube4x4x4 c5"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   DrawWidth       =   2
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "seen from left"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "seen from right"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   405
      ItemData        =   "Visual.frx":0000
      Left            =   8400
      List            =   "Visual.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton ReNew 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10560
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Value:"
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Red To Move"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   8400
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Sign 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   9000
      Shape           =   3  'Circle
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   64
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   960
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   63
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   62
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   61
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   60
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   59
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   58
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   57
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   56
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   55
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   54
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   53
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   52
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   51
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   50
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   49
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   48
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   47
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   46
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   45
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   44
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   43
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   42
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   41
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Lines 
      Index           =   0
      Visible         =   0   'False
      X1              =   240
      X2              =   1200
      Y1              =   960
      Y2              =   1800
   End
   Begin VB.Line Lines 
      Index           =   48
      X1              =   0
      X2              =   960
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Line Lines 
      Index           =   47
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   46
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   45
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   44
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   43
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   42
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   41
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   40
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   39
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   38
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   37
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   36
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   35
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   34
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   33
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   32
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   31
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   30
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   29
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   28
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   27
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   26
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   25
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   24
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   23
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   22
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   21
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   20
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   19
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   18
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   17
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   16
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   15
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   14
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   13
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   12
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   11
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   10
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   9
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   8
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   7
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   6
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   5
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   4
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   3
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   2
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Lines 
      Index           =   1
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Shape Back 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   120
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "Visual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Dim Px(4, 4, 4), Py(4, 4, 4), R(4, 4, 4), m, Flash(64) As Integer
Dim p%
Dim Still As Boolean
Dim a1, a2 As Single
Dim Record(64) As Integer
Dim la As String
Private Sub Combo1_Click()
Dim dif As Integer
dif = Combo1.ListIndex
Select Case dif
Case 0
cN = 3: c1 = 8: c2 = 7: c3 = 6
Case 1
cN = 5: c1 = 17: c2 = 13: c3 = 11
End Select
End Sub


Public Sub Draw()
Dim x, y, z, n As Integer
Dim Tx, ty, Vx, Vy, Vz, Lx As Single
Lx = 54000
Vx = 30
Vy = a1
Vz = 9.5
For z = 1 To 4
    For y = 1 To 4
       For x = 1 To 4
       R(x, y, z) = Lx / (Vx - x) / 10
      Tx = Vy * Lx / Vx
     ty = Vz * Lx / Vx
    Px(x, y, z) = (Tx - (Vy - y) * Lx / (Vx - x)) + a2
    Py(x, y, z) = 7500 + ((Vz - z) * Lx / (Vx - x) - ty)
       Next
    Next
Next
n = 0
For y = 1 To 4
For x = 1 To 4
n = n + 1
Lines(n).X1 = Px(x, y, 1)
Lines(n).X2 = Px(x, y, 4)
Lines(n).Y1 = Py(x, y, 1)
Lines(n).Y2 = Py(x, y, 4)
Next
Next
For y = 1 To 4
For z = 1 To 4
n = n + 1
Lines(n).X1 = Px(1, y, z)
Lines(n).X2 = Px(4, y, z)
Lines(n).Y1 = Py(1, y, z)
Lines(n).Y2 = Py(4, y, z)
Next
Next
For x = 1 To 4
For z = 1 To 4
n = n + 1
Lines(n).X1 = Px(x, 1, z)
Lines(n).X2 = Px(x, 4, z)
Lines(n).Y1 = Py(x, 1, z)
Lines(n).Y2 = Py(x, 4, z)
Next
Next
n = 0
For z = 1 To 4
For y = 1 To 4
For x = 1 To 4
n = n + 1
Shape1(n).Move Px(x, y, z) - R(x, y, z), Py(x, y, z) - R(x, y, z), 2 * R(x, y, z), 2 * R(x, y, z)
Next
Next
Next
End Sub

Private Sub Command1_Click()
Dim p1, p2, p3 As Integer
Still = True
stopF
Command1.Enabled = False
ReNew.Enabled = False
Remove.Enabled = False
Label1.Caption = "Thinking..."
Visual.Refresh
p = Compute - 1
p3 = p \ 16 + 1
p2 = (p Mod 16) \ 4 + 1
p1 = (p Mod 16) Mod 4 + 1
Flash(p + 1) = ToGo
PutStone p1, p2, p3
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
a1 = 8.5
a2 = 700
Combo1.ListIndex = 0
Reflect
Draw
ReNew_Click
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
stopF
If Not Button = 1 Or Still = True Then Exit Sub
Dim a, b, c, n As Integer
n = 0
For c = 1 To 4
For b = 1 To 4
For a = 1 To 4
n = n + 1
If (x - Px(a, b, c)) ^ 2 + (y - Py(a, b, c)) ^ 2 < R(a, b, c) ^ 2 Then
If Stone = 0 Then
If ((a = 1 Or a = 4) And (b = 1 Or b = 4) And (c = 1 Or c = 4)) Or ((a = 2 Or a = 3) And (b = 2 Or b = 3) And (c = 2 Or c = 3)) Then
MsgBox ("Sorry,you can't put the first stone at center or corner.")
Exit Sub
End If
End If
 If StoneLay(a, b, c) = 0 Then
   LoadStone
   PutStone a, b, c
  End If
End If
Next
Next
Next
End Sub
Public Sub PutStone(p1, p2, p3)
p = p1 + 4 * p2 + 16 * p3 - 20
Turn = ToGo
Still = True
ReNew.Enabled = True
Remove.Enabled = False
m = 0
Timer1.Enabled = True
Command1.Enabled = False
If ToGo = 1 Then
   Shape1(p).FillColor = vbRed
   Else
   Shape1(p).FillColor = vbBlue
   End If
If Win(p) = True Then
 If ToGo = 1 Then
  Label1.Caption = "Red wins!"
 Else
  Label1.Caption = "Blue wins!"
 End If
For m = 1 To 4
Flash(sL(winL, m)) = ToGo
Next
m = 0
ElseIf CheckPath = True Then
MsgBox ("Game end in a draw.")
  Else
  Label2.Caption = "Value: " & ValueX
  ValueX = Empty
  StoneLay(p1, p2, p3) = ToGo
  ToGo = -ToGo
  Stone = Stone + 1
    Record(Stone) = p
    Still = False
    Remove.Enabled = True
    Command1.Enabled = True
  If ToGo = 1 Then
 la = "Red To Play"
 Label1.Caption = la
 Sign.FillColor = vbRed
  Else
   la = "Blue To Play"
   Label1.Caption = la
   Sign.FillColor = vbBlue
   End If
   End If
End Sub


Private Sub Option1_Click()
a1 = 8.5
a2 = 700
Draw
End Sub

Private Sub Option2_Click()
a1 = -3.5
a2 = -1650
Draw
End Sub

Private Sub Remove_Click()
Dim p1, p2, p3 As Integer
stopF
ToGo = -ToGo
If ToGo = 1 Then
 la = "Red To Play"
 Label1.Caption = la
 Sign.FillColor = vbRed
Else
la = "Blue To Play"
Label1.Caption = la
Sign.FillColor = vbBlue
End If
Shape1(Record(Stone)).FillColor = &HE0E0E0
p = Record(Stone) - 1
p3 = p \ 16 + 1
p2 = (p Mod 16) \ 4 + 1
p1 = (p Mod 16) Mod 4 + 1
StoneLay(p1, p2, p3) = 0
Record(Stone) = Empty
Stone = Stone - 1
Command1.Enabled = True
If Stone = 0 Then
Remove.Enabled = False
ReNew.Enabled = False
End If
End Sub

Private Sub ReNew_Click()
Still = False
stopF
ToGo = 1
Stone = 0
For p = 1 To 64
Shape1(p).FillColor = &HE0E0E0
Next
Erase StoneLay
Erase Record
la = "Red to play"
Label1.Caption = la
Sign.FillColor = vbRed
ReNew.Enabled = False
Remove.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Timer1_Timer()
m = m + 1
For p = 1 To 64
If m Mod 2 = 1 Then
If Not Flash(p) = 0 Then Shape1(p).FillColor = vbYellow
Else
  If Flash(p) = 1 Then Shape1(p).FillColor = vbRed
  If Flash(p) = -1 Then Shape1(p).FillColor = vbBlue
End If
Next
If m > 5 Then
stopF
End If
End Sub
Private Sub stopF()
For p = 1 To 64
  If Flash(p) = 1 Then Shape1(p).FillColor = vbRed
  If Flash(p) = -1 Then Shape1(p).FillColor = vbBlue
  Flash(p) = 0
Next
Timer1.Enabled = False
End Sub
