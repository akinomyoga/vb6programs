VERSION 5.00
Object = "*\A計算ctl.vbp"
Begin VB.Form Form1 
   Caption         =   "計算機"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows の既定値
   Begin 功一関数.MathF2 MathF21 
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin 功一関数.GraphF GraphF1 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin 功一関数.MathF MathF1 
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin 功一関数.ColorF ColorF1 
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub MathF21_GotFocus()

End Sub
