VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "BLIND"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer, a

Private Sub Command1_Click()
b = Int(Rnd * 25)
If b = 0 Then b = 26
Label1.Caption = Mid(a, b, 1)
Text1.Text = ""
End Sub

Private Sub Form_Load()
a = "zxcvbnmasdfghjklqwertyuiop"
End Sub

Private Sub Text1_Change()
If Text1.Text = Label1.Caption Then
s = s + 1
Label2.Caption = s
Call Command1_Click
Else
Beep
Text1.Text = ""
End If
End Sub
