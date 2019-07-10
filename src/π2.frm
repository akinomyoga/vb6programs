VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "π計算２"
   ClientHeight    =   5115
   ClientLeft      =   4800
   ClientTop       =   4215
   ClientWidth     =   9840
   Icon            =   "π2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "マーチン２"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   4680
      ScaleHeight     =   4995
      ScaleWidth      =   4995
      TabIndex        =   15
      Top             =   0
      Width           =   5055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "　　　　　　　　　　　　　マーチンの公式　　　　　　　　　　　　　　　π = 16 arc tan (1/5) - 4 arc tan (1/239)"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Height          =   1455
      Left            =   1320
      TabIndex        =   13
      ToolTipText     =   "まだ計算していないので分かりません。"
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AC"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   840
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "π計算"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(4001) As Long, b(4001) As Long, c(4001) As Long, m1, m2

Private Sub Command1_Click()
If Label1.Caption > 0 And Label1.Caption <= 10000000 Then
Dim c As Double
bb = Label1.Caption ^ 2
d = (Label1.Caption / 2)
For aa = 0 To Label1.Caption
c = c + Int(Sqr(bb - aa ^ 2))
Next aa
c = (c + d) * 4 / bb
Command4.Caption = c
Command4.ToolTipText = Label1.Caption & "回計算して求めた数値です。"
ElseIf Label1.Caption = 0 Then
MsgBox "0を入力しても計算できません！" & Chr(13) & "数字のボタンで何か入力して下さい。"
Else
MsgBox "1000万より大きな数字を入力すると" & Chr(13) & "時間がかかり過ぎるので計算しません！"
End If
End Sub

Private Sub Command2_Click(Index As Integer)
If Label1.Caption <> 0 Then
Label1.Caption = Label1.Caption & Index
Else
Label1.Caption = Index
End If
End Sub

Private Sub Command3_Click()
Label1.Caption = 0
End Sub

Private Sub Command4_Click()
MsgBox Command4.ToolTipText
End Sub

Private Sub Command5_Click()
c1 = 5
c2 = 73 '71535 'int((int(100000桁*1/log(5))+3)/2)
For h = 1 To 2
If h = 2 Then
c1 = 239
c2 = 22 '21024
End If

For k = 1 To c2
Picture1.Circle (2500, 2500), 2500, , , k / c2 * 6.28
b(0) = 1
For i = 1 To m2 + 1
b(i) = 0
Next i
k1 = 2 * k - 1
Call TaketaWarizan(k1)
For i = 1 To k1
Call TaketaWarizan(c1)
Next i
If k Mod 2 = 0 Then
Call TaketaWarizan(-1)
End If
For i = 0 To m2
a(i) = a(i) + b(i)
If a(i) >= m1 Then
a(i) = a(i) - m1
a(i - 1) = a(i - 1) + 1
ElseIf a(i) < 0 Then
a(i) = a(i) + m1
a(i - 1) = a(i - 1) - 1
End If
Next i
Next k

Open "pi2-atn" & h & ".txt" For Output As 1
For i = 0 To 13 '3999
Print #1, a(i)
Next i
Close #1

For i = 0 To m2
If h = 1 Then
a(i) = a(i) * 16
Else
a(i) = a(i) * -4
End If
If i > 0 Then
a(i - 1) = a(i - 1) + Fix(a(i) / m1)
a(i) = a(i) Mod m1
End If
Next i

If h = 1 Then
For i = 0 To m2
c(i) = a(i)
a(i) = 0
Next i
Else
For i = 0 To m2
c(i) = c(i) + a(i)
If a(i) >= m1 Then
c(i) = c(i) - m1
c(i - 1) = c(i - 1) + 1
ElseIf c(i) < 0 Then
c(i) = c(i) + m1
c(i - 1) = c(i - 1) - 1
End If
Next i
End If

Next h

Open "pi2.txt" For Output As 1
For i = 0 To 13 '3999
Print #1, a(i)
Next i
Close #1
End Sub

Public Sub TaketaWarizan(x)
For i = 0 To m2
b(i + 1) = b(i + 1) + (b(i) Mod x) * m2
b(i) = Fix(b(i) / x)
Next i
End Sub

Private Sub Command6_Click()
Dim p(255) As Double, a1(255) As Double, b1(255) As Double, d(4)
k = 10000
n9 = 27
a1(1) = 24 * 8
mm2 = 8 ^ 2
t = 110
GoSub skip1
a1(1) = 8 * 57
mm2 = 57 ^ 2
t = 60
GoSub skip1
a1(1) = 4 * 239
mm2 = 239 ^ 2
t = 45
GoSub skip1
Open "pi1.txt" For Output As 1
Print #1, "π=" & p(1) & "."
For i = 2 To n9
a2 = p(i)
For l = 1 To 4
a3 = Int(a2 / 10)
d(l) = a2 - a3 * 10
a2 = a3
Next l
Print #1, d(4) & d(3) & d(2) & d(1)
Next i
Close #1
Exit Sub
skip1:
n = 1
loop1:
m = mm2
GoSub skip2
For i = 1 To n9
a1(i) = b1(i)
Next i
m = n
GoSub skip2
c1 = 0
For i = n9 To 1 Step -1
p(i) = p(i) + b1(i) + c1
c1 = 0
If p(i) >= k Then
p(i) = p(i) - k
c1 = 1
End If
Next i
n = n + 2
m = mm2
GoSub skip2
For i = 1 To n9
a1(i) = b1(i)
Next i
m = n
GoSub skip2
b0 = 0
For i = n9 To 1 Step -1
p(i) = p(i) - b1(i) - b0
b0 = 0
If p(i) < 0 Then
p(i) = p(i) + k
b0 = 1
End If
Next i
n = n + 2
If n < t Then GoTo loop1
Return
skip2:
r = 0
For i = 1 To n9
x = a1(i) + r * k
b1(1) = Int(x / m)
r = x - m * b1(i)
Next i
Return
End Sub

Private Sub Form_Load()
m1 = 10 ^ 8 '一つの変数の桁数
m2 = 14 '4000'使う変数の数
MsgBox Int((Int(100000 / Log(5)) + 3) / 2) & "=71535"
End Sub


Public Sub save()
For i = 0 To m2 - 1
aa = aa & b(i)
Next i
MsgBox aa
End Sub
