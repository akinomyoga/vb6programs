VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "‘ì‹…"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "‘Å‚Â"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "‘Å‚Â"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Player 1"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Player 2"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Timer Timer2 
      Left            =   5640
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   11400
      Top             =   6240
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   0
      Top             =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   5880
      X2              =   5880
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  '•s“§–¾
      BorderColor     =   &H000080FF&
      Height          =   255
      Left            =   240
      Shape           =   3  '‰~
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '•s“§–¾
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   11520
      Top             =   4920
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '•s“§–¾
      BorderColor     =   &H000000FF&
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   4920
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1200
      X2              =   10560
      Y1              =   6000
      Y2              =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub
