VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "âpíPåÍï€ë∂"
   ClientHeight    =   2370
   ClientLeft      =   6195
   ClientTop       =   780
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "ÉLÉÉÉìÉZÉã"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ï€ë∂"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ÉtÉ@ÉCÉãñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÉtÉ@ÉCÉãñºÇì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅI"
Else
If Right(Dir1.Path, 1) = "\" Then
filepath = Dir1.Path & Text1.Text & ".å˜txt"
Else
filepath = Dir1.Path & "\" & Text1.Text & ".å˜txt"
End If
Form1.filesave (filepath)
MsgBox filepath & "Ç…ï€ë∂Ç≥ÇÍÇ‹ÇµÇΩÅI"
Form4.Hide
End If
End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Drive1_Change()
On Error GoTo errhand
Dir1.Path = Drive1.Drive
Exit Sub
errhand:
MsgBox "ÉhÉâÉCÉuÇ™å©Ç¬Ç©ÇËÇ‹ÇπÇÒÅI"
End Sub

Private Sub Form_Load()
Dir1.Path = "."
End Sub
