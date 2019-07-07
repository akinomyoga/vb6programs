VERSION 5.00
Object = "*\A％メーター.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   960
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
UserControl11.ToolTipText = 0
End Sub

Private Sub Timer1_Timer()
If UserControl11.ToolTipText < 100 Then
Dim a As Integer
Dim b As Integer
Dim c As Integer
a = UserControl11.ToolTipText
b = 1
c = a + b
UserControl11.ToolTipText = c
End If
End Sub

