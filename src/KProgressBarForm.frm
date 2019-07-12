VERSION 5.00
Object = "*\AÅìÉÅÅ[É^Å[.vbp"
Begin VB.Form KProgressBarForm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   960
   End
   Begin KProgressBarProject.KProgressBar UserControl11 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _extentx        =   7435
      _extenty        =   873
   End
End
Attribute VB_Name = "KProgressBarForm"
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

