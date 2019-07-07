VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "êMçÜã@"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command1 
      Caption         =   "èIóπ"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   1320
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3360
      TabIndex        =   0
      Top             =   1800
      Width           =   495
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   3000
      Y1              =   3480
      Y2              =   3600
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   3240
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      FillStyle       =   0  'ìhÇËÇ¬Ç‘Çµ
      Height          =   255
      Left            =   2640
      Shape           =   3  'â~
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'ìhÇËÇ¬Ç‘Çµ
      Height          =   255
      Left            =   2400
      Shape           =   3  'â~
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2040
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'ìhÇËÇ¬Ç‘Çµ
      Height          =   255
      Left            =   2160
      Shape           =   3  'â~
      Top             =   1200
      Width           =   255
   End
   Begin VB.Line Line4 
      X1              =   5040
      X2              =   2760
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   1440
      Y1              =   120
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   2640
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   1560
      Y2              =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
t = 0
End Sub

Private Sub Timer1_Timer()
If t = 0 Then
Shape1.FillColor = RGB(0, 192, 192)
Shape4.FillColor = RGB(0, 0, 0)
Label2.BackColor = RGB(256, 0, 0)
ElseIf t = 10 Then
Shape1.FillColor = RGB(0, 0, 0)
Shape3.FillColor = RGB(256, 256, 0)
ElseIf t = 12 Then
Shape3.FillColor = RGB(0, 0, 0)
Shape4.FillColor = RGB(256, 0, 0)
ElseIf t = 14 Then
Label1.BackColor = RGB(0, 192, 192)
Label2.BackColor = RGB(0, 0, 0)
ElseIf t = 24 Then
Label1.BackColor = RGB(0, 0, 0)
ElseIf t = 25 Then
Label1.BackColor = RGB(0, 192, 192)
ElseIf t = 26 Then
Label1.BackColor = RGB(0, 0, 0)
ElseIf t = 27 Then
Label1.BackColor = RGB(0, 192, 192)
ElseIf t = 28 Then
Label1.BackColor = RGB(0, 0, 0)
ElseIf t = 29 Then
Label1.BackColor = RGB(0, 192, 192)
End If
If t = 31 Then
t = 0
Label2.BackColor = RGB(256, 0, 0)
Label1.BackColor = RGB(0, 0, 0)
Else
t = t + 1
End If
End Sub
