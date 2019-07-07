VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "Tetris"
   ClientHeight    =   2610
   ClientLeft      =   6390
   ClientTop       =   3765
   ClientWidth     =   2085
   Icon            =   "tetris2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(20, 9, 1) As Byte, b(3, 1) As Integer, c(20, 9, 1) As Byte
Private Sub Command1_Click()
Randomize
Call Tugi
Timer1.Interval = 200
Command1.Enabled = False
End Sub

Private Sub left1()
For b1 = 0 To 3
If b(b1, 0) < 1 Then Exit Sub
If a(b(b1, 1), b(b1, 0) - 1, 0) = 1 Then Exit Sub
Next b1
For b1 = 0 To 3
b(b1, 0) = b(b1, 0) - 1
Next b1
Call hyoji
End Sub

Private Sub turn1()
Dim b2(1 To 3, 1)
For b1 = 1 To 3
b3 = b(b1, 0) - b(0, 0)
b4 = b(b1, 1) - b(0, 1)
b2(b1, 0) = b4 + b(0, 0)
b2(b1, 1) = -b3 + b(0, 1)
If b2(b1, 0) < 0 Or 9 < b2(b1, 0) Or b2(b1, 1) < 0 Or 20 < b2(b1, 1) Then Exit Sub
If a(b2(b1, 1), b2(b1, 0), 0) = 1 Then Exit Sub
Next b1
For b1 = 1 To 3
b(b1, 0) = b2(b1, 0)
b(b1, 1) = b2(b1, 1)
Next b1
Call hyoji
End Sub

Private Sub right1()
For b1 = 0 To 3
If b(b1, 0) > 8 Then Exit Sub
If a(b(b1, 1), b(b1, 0) + 1, 0) = 1 Then Exit Sub
Next b1
For b1 = 0 To 3
b(b1, 0) = b(b1, 0) + 1
Next b1
Call hyoji
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 37
Call left1
Case 32
Call turn1
Case 39
Call right1
End Select
End Sub

Private Sub Timer1_Timer()
For b1 = 0 To 3
If b(b1, 1) = 0 Then GoTo Tugi
If a(b(b1, 1) - 1, b(b1, 0), 0) = 1 Then GoTo Tugi
Next b1
For b1 = 0 To 3
b(b1, 1) = b(b1, 1) - 1
Next b1
Call hyoji
Exit Sub
Tugi:
For b1 = 0 To 3
a(b(b1, 1), b(b1, 0), 0) = 1
Next b1
Dim a3 As Integer
For a1 = 0 To 20
For a2 = 0 To 9
a3 = a3 + a(a1, a2, 0)
Next a2
If a3 = 10 Then
 For a4 = 0 To 9
 For a3 = a1 To 19
 a(a3, a4, 0) = a(a3 + 1, a4, 0)
 Next a3
 a(20, a4, 0) = 0
 Next a4
 Label1.Caption = Label1.Caption + 1
 a1 = a1 - 1
End If
a3 = 0
Next a1
Call Tugi
Call hyoji
End Sub

Private Sub Tugi()
num = Int(Rnd * 7)
If num = 7 Then num = 0
Select Case num
Case 0
Call sett(4, 20, 3, 20, 5, 20, 6, 20)
Case 1
Call sett(5, 20, 4, 20, 4, 19, 5, 19)
Case 2
Call sett(5, 20, 4, 20, 6, 20, 5, 19)
Case 3
Call sett(5, 20, 4, 20, 6, 20, 4, 19)
Case 4
Call sett(5, 20, 4, 20, 6, 20, 6, 19)
Case 5
Call sett(5, 20, 4, 20, 5, 19, 6, 19)
Case 6
Call sett(5, 20, 6, 20, 5, 19, 4, 19)
End Select
End Sub

Private Sub sett(a1, a2, a3, a4, a5, a6, a7, a8)
b(0, 0) = a1: b(0, 1) = a2
b(1, 0) = a3: b(1, 1) = a4
b(2, 0) = a5: b(2, 1) = a6
b(3, 0) = a7: b(3, 1) = a8
For b1 = 0 To 3
If a(b(b1, 1), b(b1, 0), 0) = 1 Then
MsgBox "GameOver"
Timer1.Interval = 0
Command1.Enabled = True
Label1.Caption = 0
Form2.Cls
For a1 = 0 To 20
For a2 = 0 To 9
a(a1, a2, 0) = 0
a(a1, a2, 1) = 0
c(a1, a2, 0) = 0
c(a1, a2, 1) = 0
Next a2
Next a1
Call sett(0, 0, 0, 0, 0, 0, 0, 0)
End If
Next b1
End Sub

Private Sub hyoji()
For b1 = 0 To 3
c(b(b1, 1), b(b1, 0), 1) = 1
Next b1
For a1 = 0 To 20
For a2 = 0 To 9
If a(a1, a2, 0) = 1 Then c(a1, a2, 1) = 1
Next a2
Next a1
For a1 = 0 To 20
For a2 = 0 To 9
If c(a1, a2, 1) > c(a1, a2, 0) Then
Call poin(a2, a1, &H0)
c(a1, a2, 0) = 1
ElseIf c(a1, a2, 1) < c(a1, a2, 0) Then
Call poin(a2, a1, &HFFFFFF)
c(a1, a2, 0) = 0
End If
c(a1, a2, 1) = 0
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
