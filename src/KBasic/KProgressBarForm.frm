VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#5.0#0"; "KBasic.ocx"
Begin VB.Form KProgressBarForm 
   Caption         =   "Test KProgressBar"
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
      Left            =   240
      Top             =   720
   End
   Begin KBasic.KProgressBar UserControl11 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
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
    If UserControl11.Value < 100 Then
        UserControl11.Value = UserControl11.Value + 1
    Else
        UserControl11.Value = 0
    End If
End Sub

