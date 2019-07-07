VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Graph"
   ClientHeight    =   7905
   ClientLeft      =   2220
   ClientTop       =   765
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10500
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   9840
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      TabIndex        =   4
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "y=       x+"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      X1              =   3900
      X2              =   3900
      Y1              =   7800
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      X1              =   7800
      X2              =   0
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'ïsìßñæ
      BorderColor     =   &H00808080&
      FillColor       =   &H0000FF00&
      FillStyle       =   6  '∏€Ω
      Height          =   7815
      Left            =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Cls
On Error GoTo ERR
ya = Text1.Text * 3900 - Text2.Text * 120 + 3900
yb = Text1.Text * -3900 - Text2.Text * 120 + 3900
Line (0, ya)-(7800, yb), RGB(0, 0, 0)
Exit Sub
ERR: MsgBox "êÆêîÇ‹ÇΩÇÕè¨êîÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
End Sub
