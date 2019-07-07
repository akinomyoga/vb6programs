VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   4545
   ClientTop       =   3015
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "vvvv"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "G"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(9, 20, 2) As Byte


Private Sub Timer1_Timer()
For a1 = 0 To 9
For a2 = 0 To 20
If a(a1, a2, 2) = 1 Then
If a2 > 0 Then
If a(a1, a2 - 1, 0) = 0 Then
a(a1, a2 - 1, 1) = 1
Else
GoTo Tugi
End If
Else
GoTo Tugi
End If
End If
Next a2
Next a1

For a1 = 0 To 9
For a2 = 0 To 20
a(a1, a2, 2) = a(a1, a2, 1)
a(a1, a2, 1) = 0
Next a2
Next a1

Call hyoji
Exit Sub
Tugi:
For a1 = 0 To 9
For a2 = 0 To 20
a(a1, a2, 1) = 0
Next a2
Next a1
Call Tugi
End Sub
Private Sub hyoji()
For a1 = 0 To 9
For a2 = 0 To 20
If a(a1, a2, 0) = 1 Or a(a1, a2, 2) = 1 Then
Call poin(a1, a2, &H0)
Else
Call poin(a1, a2, &HFFFFFF)
End If
Next a2
Next a1
End Sub

Private Sub poin(a1, a2, col)
For x = 0 To 8
For y = 0 To 8
PSet (x * 15 + a1 * 120, 2400 + y * 15 - a2 * 120), col
Next y
Next x
End Sub

Private Sub Tugi()
Dim a3 As Integer
For a2 = 0 To 20
For a1 = 0 To 9
If a(a1, a2, 2) = 1 Then
a(a1, a2, 0) = 1
a(a1, a2, 2) = 0
End If
If a(a1, a2, 0) = 1 Then a3 = a3 + 1
Next a1
If a3 = 10 Then
For a3 = 0 To 9
For a4 = a2 To 19
a(a3, a4, 0) = a(a3, a4 + 1, 0)
Next a4
a(a3, 20, 0) = 0
Next a3
End If
Next a2
num = Int(Rnd * 7): If num = 7 Then num = 6
Select Case num
Case 0
For a1 = -1 To 2
Call sett(a1, 0)
Next a1
Case 1
For a1 = 0 To 1
For a2 = 0 To 1
Call sett(a1, a2)
Next a2
Next a1
Case 2
For a1 = 0 To 1
For a2 = 0 To 1
Call sett(a1 + a2, a2)
Next a2
Next a1
Case 3
For a1 = 0 To 1
For a2 = 0 To 1
Call sett(a1 - a2, a2)
Next a2
Next a1
Case 4
For a1 = -1 To 1
Call sett(a1, 0)
Next a1
Call sett(-1, 1)
Case 5
For a1 = -1 To 1
Call sett(a1, 0)
Next a1
Call sett(0, 1)
Case 6
For a1 = -1 To 1
Call sett(a1, 0)
Next a1
Call sett(1, 1)
End Select
End Sub

Private Sub sett(a1, a2)
aa1 = a1 + 4
aa2 = 20 - a2
If a(aa1, aa2, 0) = 1 Then
MsgBox "GameOver"
Timer1.Interval = 0
Command1.Enabled = True
For a1 = 0 To 9
For a2 = 0 To 20
a(a1, a2, 0) = 0
a(a1, a2, 1) = 0
Next a2
Next a1
Else
a(aa1, aa2, 2) = 1
End If
End Sub
