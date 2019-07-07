VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "UNKNOWN"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command6 
      Caption         =   "ïœä∑"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   31
      Text            =   "ñ≥ëË"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "end"
      Height          =   375
      Left            =   1560
      TabIndex        =   30
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "save                 "
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ent"
      Height          =   375
      Left            =   1920
      TabIndex        =   28
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "spc"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      ScrollBars      =   3  'óºï˚
      TabIndex        =   26
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "z"
      Height          =   375
      Index           =   25
      Left            =   1680
      TabIndex        =   25
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "y"
      Height          =   375
      Index           =   24
      Left            =   1320
      TabIndex        =   24
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      Height          =   375
      Index           =   23
      Left            =   2040
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "w"
      Height          =   375
      Index           =   22
      Left            =   1680
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "v"
      Height          =   375
      Index           =   21
      Left            =   1320
      TabIndex        =   21
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "u"
      Height          =   375
      Index           =   20
      Left            =   2040
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "t"
      Height          =   375
      Index           =   19
      Left            =   1680
      TabIndex        =   19
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "s"
      Height          =   375
      Index           =   18
      Left            =   1320
      TabIndex        =   18
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "r"
      Height          =   375
      Index           =   17
      Left            =   2040
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "q"
      Height          =   375
      Index           =   16
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "p"
      Height          =   375
      Index           =   15
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "o"
      Height          =   375
      Index           =   14
      Left            =   840
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "n"
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "m"
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "l"
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "k"
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "j"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "i"
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "h"
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "g"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "f"
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "d"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "c"
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "b"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "a"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "HGSçsèëëÃ"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   2775
      Left            =   2520
      TabIndex        =   33
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
Text1.Text = Text1.Text & "@"
Case 1
Text1.Text = Text1.Text & "A"
Case 2
Text1.Text = Text1.Text & "B"
Case 3
Text1.Text = Text1.Text & "C"
Case 4
Text1.Text = Text1.Text & "D"
Case 5
Text1.Text = Text1.Text & "E"
Case 6
Text1.Text = Text1.Text & "F"
Case 7
Text1.Text = Text1.Text & "G"
Case 8
Text1.Text = Text1.Text & "H"
Case 9
Text1.Text = Text1.Text & "I"
Case 10
Text1.Text = Text1.Text & "J"
Case 11
Text1.Text = Text1.Text & "K"
Case 12
Text1.Text = Text1.Text & "L"
Case 13
Text1.Text = Text1.Text & "M"
Case 14
Text1.Text = Text1.Text & "N"
Case 15
Text1.Text = Text1.Text & "O"
Case 16
Text1.Text = Text1.Text & "P"
Case 17
Text1.Text = Text1.Text & "Q"
Case 18
Text1.Text = Text1.Text & "R"
Case 19
Text1.Text = Text1.Text & "S"
Case 20
Text1.Text = Text1.Text & "T"
Case 21
Text1.Text = Text1.Text & "U"
Case 22
Text1.Text = Text1.Text & "V"
Case 23
Text1.Text = Text1.Text & "W"
Case 24
Text1.Text = Text1.Text & "X"
Case 25
Text1.Text = Text1.Text & "Y"
End Select
Label1.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & "   "
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & Chr(13)
End Sub

Private Sub Command4_Click()
On Error GoTo errh
Open aa & Text2.Text & ".txt" For Output As 1
Print #1, Text1.Text
Close #1
MsgBox aa & "Ç…ï€ë∂Ç≥ÇÍÇ‹ÇµÇΩÅB"
Exit Sub
errh:
MsgBox "ï€ë∂Ç≈Ç´Ç‹ÇπÇÒÇ≈ÇµÇΩÅB"
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
a = Text1.Text
Text1.Text = ""
For b = 1 To Len(a)
Select Case Mid(a, b, 1)
Case "a"
Call Command1_Click(0)
Case "b"
Call Command1_Click(1)
Case "c"
Call Command1_Click(2)
Case "d"
Call Command1_Click(3)
Case "e"
Call Command1_Click(4)
Case "f"
Call Command1_Click(5)
Case "g"
Call Command1_Click(6)
Case "h"
Call Command1_Click(7)
Case "i"
Call Command1_Click(8)
Case "j"
Call Command1_Click(9)
Case "k"
Call Command1_Click(10)
Case "l"
Call Command1_Click(11)
Case "m"
Call Command1_Click(12)
Case "n"
Call Command1_Click(13)
Case "o"
Call Command1_Click(14)
Case "p"
Call Command1_Click(15)
Case "q"
Call Command1_Click(16)
Case "r"
Call Command1_Click(17)
Case "s"
Call Command1_Click(18)
Case "t"
Call Command1_Click(19)
Case "u"
Call Command1_Click(20)
Case "v"
Call Command1_Click(21)
Case "w"
Call Command1_Click(22)
Case "x"
Call Command1_Click(23)
Case "y"
Call Command1_Click(24)
Case "z"
Call Command1_Click(25)
Case Else
Text1.Text = Text1.Text + Mid(a, b, 1)
End Select
Next b
Label1.Caption = Text1.Text
End Sub

Private Sub Form_Load()
aa = CurDir & "\"
End Sub

