VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "game"
   ClientHeight    =   4875
   ClientLeft      =   2070
   ClientTop       =   2340
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7605
   Begin VB.CommandButton Command5 
      Caption         =   "A"
      Height          =   300
      Left            =   6480
      TabIndex        =   4
      Top             =   3600
      Width           =   735
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
      Left            =   6720
      TabIndex        =   3
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6960
      TabIndex        =   2
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
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
      Left            =   6720
      TabIndex        =   0
      Top             =   2880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  '塗りつぶし
      Height          =   255
      Left            =   0
      Shape           =   3  '円
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "rpg1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(15, 15), b(5, 5, 15, 15) As Integer, c(15, 15), d(5, 5, 15, 15), x As Integer, y As Integer, xb As Integer, yb As Integer
Private Sub IDOU()
Shape1.Top = y * 240
Shape1.Left = x * 240
End Sub

Private Sub Command1_Click()
If y > 0 Then
If a(x, y - 1) = -1 Then
Else
y = y - 1
Call IDOU
End If
End If
End Sub

Private Sub Command2_Click()
If x > 0 Then
If a(x - 1, y) = -1 Then
Else
x = x - 1
Call IDOU
End If
End If
End Sub

Private Sub Command3_Click()
If x < 15 Then
If a(x + 1, y) = -1 Then
Else
x = x + 1
Call IDOU
End If
End If
End Sub

Private Sub Command4_Click()
If y < 15 Then
If a(x, y + 1) = -1 Then
Else
y = y + 1
Call IDOU
End If
End If
End Sub

Private Sub Command5_Click()
MsgBox c(x, y)
End Sub

Private Sub Form_Load()
b(0, 0, 0, 0) = -1 '進めないところを設定する。始めの二つは背景の位置　後の二つは画面中の位置。-1は「進めない」
d(0, 0, 0, 0) = "誰かが居る！"
x = 0
y = 0
Call HaikeiSet
End Sub

Private Sub HaikeiSet()
For ba = 0 To 15
For bb = 0 To 15
a(ba, bb) = b(xb, yb, ba, bb)
c(ba, bb) = d(xb, yb, ba, bb)
Next bb
Next ba
End Sub
