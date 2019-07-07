VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "H•¨˜A½"
   ClientHeight    =   1785
   ClientLeft      =   6690
   ClientTop       =   3750
   ClientWidth     =   1560
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   840
      Max             =   30
      Min             =   5
      TabIndex        =   2
      Top             =   480
      Value           =   5
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa(1, 501, 501, 2) As Integer, bb(2) As Integer, cc(2) As Integer, dd(1), oo(2)

Private Sub Command1_Click()
Timer1.Interval = 1
End Sub

Private Sub Command2_Click()
Timer1.Interval = 0
End Sub

Private Sub Form_Load()
Randomize
dd(0) = 50 '‰¡
dd(1) = 50 'c
bb(0) = 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Select Case Shift
Case False
aa(0, X / 15, Y / 15, 1) = 2
aa(1, X / 15, Y / 15, 1) = 2
Case True
aa(0, X / 15, Y / 15, 1) = 0
aa(1, X / 15, Y / 15, 1) = 0
End Select
ElseIf Button = 4 Then
MsgBox X / 15 & ":" & Y / 15 & Chr(13) & aa(0, X / 15, Y / 15, 0) & "," & aa(0, X / 15, Y / 15, 1) & "," & aa(0, X / 15, Y / 15, 2)
ElseIf Button = 2 Then
Select Case Shift
Case False
aa(0, X / 15, Y / 15, 2) = 2
aa(1, X / 15, Y / 15, 2) = 2
Case True
aa(0, X / 15, Y / 15, 2) = 0
aa(1, X / 15, Y / 15, 2) = 0
End Select
End If
End Sub

Private Sub Timer1_Timer()
Dim d(2) As Integer, e As Integer
For a = 1 To dd(0)
For b = 1 To dd(1)
For c = 0 To 2
aa(1, a, b, c) = aa(0, a, b, c)
Next c
Next b
Next a
For a = 1 To dd(0)
For b = 1 To dd(1)
For c = 0 To 2
d(c) = aa(1, a - 1, b - 1, c) + aa(1, a, b - 1, c) + aa(1, a + 1, b - 1, c) + aa(1, a - 1, b, c) + aa(1, a, b, c) + aa(1, a + 1, b, c) + aa(1, a - 1, b + 1, c) + aa(1, a, b + 1, c) + aa(1, a + 1, b + 1, c)
Next c
aa(0, a, b, 0) = aa(0, a, b, 0) + bb(0) - aa(0, a, b, 1)
If aa(0, a, b, 0) > 6 Then aa(0, a, b, 0) = 6
If aa(0, a, b, 0) < 0 Then aa(0, a, b, 0) = 0
If aa(1, a, b, 1) <= aa(1, a, b, 0) And d(1) < 19 And d(1) > 5 Then
If d(1) < 13 Then
e = 2
Else
e = 1 + Int(Rnd * 8 / (12 - d(1)))
End If
End If
aa(0, a, b, 1) = aa(0, a, b, 1) - 1 - aa(0, a, b, 2) + e
If aa(0, a, b, 1) < 0 Then aa(0, a, b, 1) = 0
e = 0
If d(1) > 5 And d(2) = 2 Then e = 2
aa(0, a, b, 2) = aa(0, a, b, 2) - 1 + e
If aa(0, a, b, 2) < 0 Then aa(0, a, b, 2) = 0
Form1.PSet (a * 15, b * 15), RGB(255 / 1 * aa(0, a, b, 2), 255 / 6 * aa(0, a, b, 0), 255 / 3 * aa(0, a, b, 1))
ff = ff + aa(0, a, b, 0)
gg = gg + aa(0, a, b, 1)
hh = hh + aa(0, a, b, 2)
Next b
Next a
cc(2) = cc(2) + 1
If cc(2) > 100 Then cc(2) = 0
ia = 1500 - cc(2) * 15
ib = 1440 / dd(0) / dd(1)
ic = Picture1.Height - 195
Picture1.Line (ia, ic)-(ia, 0), RGB(0, 0, 0)
Picture1.Line (ia, ic - ib / 10 * ff)-(ia + 15, ic - ib / 10 * oo(0)), RGB(0, 255, 0)
Picture1.Line (ia, ic - ib / 10 * gg)-(ia + 15, ic - ib / 10 * oo(1)), RGB(0, 0, 255)
Picture1.Line (ia, ic - ib / 10 * hh)-(ia + 15, ic - ib / 10 * oo(2)), RGB(255, 0, 0)
oo(0) = ff
oo(1) = gg
oo(2) = hh
cc(0) = cc(0) + 1
If cc(0) > HScroll1.Value Then
cc(0) = 0
cc(1) = cc(1) + 1
If cc(1) > 3 Then cc(1) = 0
Select Case cc(1)
Case 0
bb(0) = 1
Case 2
bb(0) = 1
Case 1
bb(0) = 2
Case 3
bb(0) = 0
End Select
End If
Form1.Caption = cc(1) & " : " & cc(0)
End Sub
