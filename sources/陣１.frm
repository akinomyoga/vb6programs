VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "Ç‹Ç∑ÇÃêî"
   ClientHeight    =   1590
   ClientLeft      =   4650
   ClientTop       =   4215
   ClientWidth     =   3990
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   0
      Max             =   30
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   30
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ìKóp"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "30Å~30,çáåv900Ç‹Ç∑"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Call Command2_Click
End Sub

Private Sub Command2_Click()
Form1.ccnt = HScroll1.Value
Call Form1.DRAWCELL
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Form2.Hide
End Sub

Private Sub HScroll1_Change()
h1 = HScroll1.Value
Label1.Caption = h1 & "Å~" & h1 & ",çáåv" & h1 ^ 2 & "Ç‹Ç∑"
Command2.Enabled = True
End Sub
