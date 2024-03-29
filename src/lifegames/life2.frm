VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Life Game-0回計算"
   ClientHeight    =   4335
   ClientLeft      =   6360
   ClientTop       =   5025
   ClientWidth     =   3615
   FillColor       =   &H0000FFFF&
   Icon            =   "life2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3615
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  '塗りつぶし
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   1500
      TabIndex        =   28
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command18 
         Caption         =   "0回計算に戻す"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton Command17 
         Caption         =   "これについて"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         MaskColor       =   &H8000000F&
         TabIndex        =   25
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         Caption         =   "35%発生"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         ToolTipText     =   "表示されている確率でランダムに点をセットしていきます。"
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   20
         TabIndex        =   23
         Top             =   1080
         Value           =   7
         Width           =   855
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   960
         Max             =   30
         Min             =   1
         TabIndex        =   20
         Top             =   1800
         Value           =   1
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Height          =   255
         Left            =   600
         TabIndex        =   16
         ToolTipText     =   "右下に移動"
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "END"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         ToolTipText     =   "このプログラムを終了します。右上の×ボタンでも出来ます。"
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "消去"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         ToolTipText     =   "また始めからやり直したい時などに押すと始めの状態に戻ります。"
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "□"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   5.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         ToolTipText     =   "移動する青枠を見て点を打ちたい時に青枠を表示する。"
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Height          =   255
         Left            =   960
         TabIndex        =   11
         ToolTipText     =   "移動する青枠がじゃまになった時などに隠す。"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         ToolTipText     =   "右に移動"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "下に移動"
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "○"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "青枠の真ん中に点をつける。"
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "fastest"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "最も速く計算します。"
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Height          =   255
         Left            =   600
         TabIndex        =   17
         ToolTipText     =   "右上に移動"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "上に移動"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command13 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "左下に移動"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "fast"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "速めに計算します。"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "左に移動"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "左上に移動"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "slow"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "ゆっくりと計算します。"
         Top             =   2040
         Width           =   735
      End
      Begin VB.OptionButton Option0 
         Caption         =   "stop"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "計算しません。休みます。"
         Top             =   1800
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "mid."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "普通の速さで計算します。"
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   19
         ToolTipText     =   "間違えてしまった時などに要らない点を消す。"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command16 
         Caption         =   "表示"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "他のWindowがかぶってしまって画面が消えてしまった時などに押すとまた表示されます。"
         Top             =   3720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "詳細に計算"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   6.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         Height          =   1815
         Left            =   105
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "1度に1回計算"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   6.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   21
         ToolTipText     =   "一度に表示されている回数だけ計算してそれから表示します。"
         Top             =   2040
         Width           =   930
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   75
      Left            =   600
      Top             =   600
      Width           =   75
   End
   Begin VB.Shape Shape1 
      Height          =   1530
      Left            =   -15
      Top             =   -15
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a(101, 101) As Boolean, b(101, 101) As Boolean, V As Integer
Dim ao(99) As Integer, bo As Integer, ap(99) As Integer, aq(99) As Integer
Dim ar(99) As Integer

Private Sub DRAWING()
DrawWidth = 1
Cls
For X = 0 To 99
For Y = 0 To 99
If a(X + 1, Y + 1) = True Then
PSet (X * 15, Y * 15), RGB(255, 0, 0)
End If
Next Y
Next X
End Sub

Private Sub SQUARE(X As Integer, Y As Integer)
If X >= 0 And X <= 99 Then
If Y >= 0 And Y <= 99 Then
Shape2.Top = (Y - 2) * 15
Shape2.Left = (X - 2) * 15
End If
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = True Then
Timer2.Interval = Timer1.Interval
Timer1.Interval = 0
For bo2 = 0 To 99
ao(bo2) = 0
Next bo2
bo = 99
Else
Timer1.Interval = Timer2.Interval
Timer2.Interval = 0
End If
End Sub

Private Sub Command1_Click()
If Shape2.Left > -30 Then Shape2.Left = Shape2.Left - 15
Call DRAWING
End Sub

Private Sub Command10_Click()
If Shape2.Left > -30 Then Shape2.Left = Shape2.Left - 15
If Shape2.Top > -30 Then Shape2.Top = Shape2.Top - 15
Call DRAWING
End Sub

Private Sub Command11_Click()
If Shape2.Top < 1455 Then Shape2.Top = Shape2.Top + 15
If Shape2.Left < 1455 Then Shape2.Left = Shape2.Left + 15
Call DRAWING
End Sub

Private Sub Command12_Click()
If Shape2.Top > -30 Then Shape2.Top = Shape2.Top - 15
If Shape2.Left < 1455 Then Shape2.Left = Shape2.Left + 15
Call DRAWING
End Sub

Private Sub Command13_Click()
If Shape2.Top < 1455 Then Shape2.Top = Shape2.Top + 15
If Shape2.Left > -30 Then Shape2.Left = Shape2.Left - 15
Call DRAWING
End Sub

Private Sub Command14_Click()
X = Shape2.Left + 30
Y = Shape2.Top + 30
If X >= 0 And Y >= 0 Then
If X < 1500 And Y < 1500 Then
a(X / 15 + 1, Y / 15 + 1) = False
b(X / 15 + 1, Y / 15 + 1) = False
Call DRAWING
End If
End If
End Sub

Private Sub Command15_Click()
Randomize
For X = 0 To 99
For Y = 0 To 99
z = Rnd
If z <= HScroll2.Value / 20 Then
a(X + 1, Y + 1) = True
b(X + 1, Y + 1) = True
Else
a(X + 1, Y + 1) = False
b(X + 1, Y + 1) = False
End If
Next Y
Next X
Call DRAWING
End Sub

Private Sub Command16_Click()
Call DRAWING
End Sub

Private Sub Command17_Click()
MsgBox "作った人　村瀬功一" & Chr(13) & "無断で改良禁止"
End Sub

Private Sub Command18_Click()
V = 0
Form1.Caption = "Life Game-" & V & "回計算"
End Sub

Private Sub Command2_Click()
If Shape2.Top > -30 Then Shape2.Top = Shape2.Top - 15
Call DRAWING
End Sub

Private Sub Command3_Click()
X = Shape2.Left + 30
Y = Shape2.Top + 30
If X >= 0 And Y >= 0 Then
If X < 1500 And Y < 1500 Then
a(X / 15 + 1, Y / 15 + 1) = True
b(X / 15 + 1, Y / 15 + 1) = True
Call DRAWING
End If
End If
End Sub

Private Sub Command4_Click()
If Shape2.Top < 1455 Then Shape2.Top = Shape2.Top + 15
Call DRAWING
End Sub

Private Sub Command5_Click()
If Shape2.Left < 1455 Then Shape2.Left = Shape2.Left + 15
Call DRAWING
End Sub

Private Sub Command6_Click()
Shape2.Visible = False
Call DRAWING
End Sub

Private Sub Command7_Click()
Shape2.Visible = True
Call DRAWING
End Sub

Private Sub Command8_Click()
For X = 0 To 99
For Y = 0 To 99
a(X + 1, Y + 1) = False
b(X + 1, Y + 1) = False
Next Y
Next X
Call DRAWING
End Sub

Private Sub Command9_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X >= 0 And Y >= 0 Then
If X < 1500 And Y < 1500 Then
a(X / 15 + 1, Y / 15 + 1) = True
b(X / 15 + 1, Y / 15 + 1) = True
Call SQUARE(X / 15, Y / 15)
Call DRAWING
End If
End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And Y >= 0 Then
If X < 1500 And Y < 1500 Then
If Button = 1 Then
a(X / 15 + 1, Y / 15 + 1) = True
b(X / 15 + 1, Y / 15 + 1) = True
Call DRAWING
End If
Call SQUARE(X / 15, Y / 15)
End If
End If
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "1度に" & HScroll1.Value & "回計算"
End Sub

Private Sub HScroll2_Change()
Command15.Caption = HScroll2.Value * 5 & "%発生"
End Sub

Private Sub Option0_Click()
SETTIMER (0)
End Sub

Private Sub Option1_Click()
SETTIMER (1000)
End Sub

Private Sub Option2_Click()
SETTIMER (500)
End Sub

Private Sub Option3_Click()
SETTIMER (100)
End Sub

Private Sub Option4_Click()
SETTIMER (1)
End Sub

Private Sub Timer1_Timer()
For c = 1 To HScroll1.Value
Dim z As Integer
For X = 1 To 100
For Y = 1 To 100
z = 0
If a(X - 1, Y) = True Then z = z + 1
If a(X - 1, Y - 1) = True Then z = z + 1
If a(X - 1, Y + 1) = True Then z = z + 1
If a(X + 1, Y) = True Then z = z + 1
If a(X + 1, Y - 1) = True Then z = z + 1
If a(X + 1, Y + 1) = True Then z = z + 1
If a(X, Y - 1) = True Then z = z + 1
If a(X, Y + 1) = True Then z = z + 1
If z = 3 Then
b(X, Y) = True
If a(X, Y) = False Then PSet (15 * X - 15, 15 * Y - 15), RGB(255, 0, 0)
ElseIf z > 3 Or z < 2 Then
b(X, Y) = False
If a(X, Y) = True Then PSet (15 * X - 15, 15 * Y - 15), RGB(0, 255, 0)
End If
Next Y
Next X
For X = 1 To 100
For Y = 1 To 100
a(X, Y) = b(X, Y)
Next Y
Next X
Next c
V = V + HScroll1.Value
Form1.Caption = "Life Game-" & V & "回計算"
End Sub

Private Sub Timer2_Timer()
Dim z As Integer, o As Integer
For c = 1 To HScroll1.Value
o = 0 '全体の数  ここに数値を
p = 0 '生き延びた数
q = 0 '生まれた数
r = 0 '死んだ数
For X = 1 To 100
For Y = 1 To 100
z = 0
If a(X - 1, Y) = True Then z = z + 1
If a(X - 1, Y - 1) = True Then z = z + 1
If a(X - 1, Y + 1) = True Then z = z + 1
If a(X + 1, Y) = True Then z = z + 1
If a(X + 1, Y - 1) = True Then z = z + 1
If a(X + 1, Y + 1) = True Then z = z + 1
If a(X, Y - 1) = True Then z = z + 1
If a(X, Y + 1) = True Then z = z + 1
If z = 3 Then
b(X, Y) = True
o = o + 1
If a(X, Y) = True Then
p = p + 1
Else
q = q + 1
PSet (15 * X - 15, 15 * Y - 15), RGB(255, 0, 0)
End If
ElseIf z = 2 Then
If b(X, Y) = True Then
o = o + 1
p = p + 1
End If
Else
b(X, Y) = False
If a(X, Y) = True Then
r = r + 1
PSet (15 * X - 15, 15 * Y - 15), RGB(0, 255, 0)
End If
End If
Next Y
Next X
For X = 1 To 100
For Y = 1 To 100
a(X, Y) = b(X, Y)
Next Y
Next X
bo = bo + 1
If bo = 100 Then bo = 0
ao(bo) = o 'ここに数値を
ap(bo) = p
aq(bo) = q
ar(bo) = r
Next c

Call DRAWGRAPH
V = V + HScroll1.Value
Form1.Caption = "Life Game-" & V & "回計算"
End Sub

Public Sub SETTIMER(inter As Integer)
If Check1.Value = False Then
Timer1.Interval = inter
Else
Timer2.Interval = inter
End If
End Sub

Public Sub DRAWGRAPH()
Picture1.Cls
Dim c As Integer, oao As Integer, pap As Integer, qaq As Integer 'ここに数値を
Dim rar As Integer
c = bo + 1
If c = 100 Then c = 0
oao = ao(c) 'ここに数値を
pap = ap(c)
qaq = aq(c)
rar = ar(c)
For bo0 = 0 To 99
c = bo0 + bo + 1
If c > 99 Then c = c - 100
Picture1.Line (bo0 * 15 + 30, 2700 - ao(c) / 4)-(bo0 * 15 + 15, 2700 - oao / 4), RGB(0, 0, 255)
Picture1.Line (bo0 * 15 + 30, 2700 - ap(c) / 4)-(bo0 * 15 + 15, 2700 - pap / 4), RGB(0, 255, 0)
Picture1.Line (bo0 * 15 + 30, 900 - aq(c) / 2)-(bo0 * 15 + 15, 900 - qaq / 2), RGB(200, 200, 0)
Picture1.Line (bo0 * 15 + 30, 900 + ar(c) / 2)-(bo0 * 15 + 15, 900 + rar / 2), RGB(255, 0, 0)
Picture1.Line (bo0 * 15 + 30, 900 - (aq(c) - ar(c)))-(bo0 * 15 + 15, 900 - (qaq - rar)), RGB(0, 150, 150)
oao = ao(c) 'ここに数値を
pap = ap(c)
qaq = aq(c)
rar = ar(c)
Next bo0
End Sub
