VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   4365
   ClientTop       =   2190
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   7500
   Begin VB.OptionButton Option4 
      Caption         =   "Å‘¬"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   7560
      Width           =   855
   End
   Begin VB.OptionButton Option0 
      Caption         =   "ÃŽ~"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   7560
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   7560
   End
   Begin VB.OptionButton Option3 
      Caption         =   "’x‚¢"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "’†ˆÊ"
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   7560
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "‘¬‚¢"
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   7800
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(249, 249) As Boolean, b(249, 249) As Boolean

Private Sub DRAWING()
DrawWidth = 2
Cls
For X = 0 To 249
For Y = 0 To 249
If a(X, Y) = True Then
PSet (X * 30, Y * 30), RGB(0, 255, 0)
End If
Next Y
Next X
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X >= 0 And Y >= 0 Then
If X < 7500 And Y < 7500 Then
a(Int(X / 30), Int(Y / 30)) = True
b(Int(X / 30), Int(Y / 30)) = True
Call DRAWING
End If
End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X >= 0 And Y >= 0 Then
If X < 7500 And Y < 7500 Then
a(Int(X / 30), Int(Y / 30)) = True
b(Int(X / 30), Int(Y / 30)) = True
Call DRAWING
End If
End If
End If
End Sub

Private Sub Option0_Click()
Timer1.Interval = 0
End Sub

Private Sub Option1_Click()
Timer1.Interval = 100
End Sub

Private Sub Option2_Click()
Timer1.Interval = 500
End Sub

Private Sub Option3_Click()
Timer1.Interval = 1000
End Sub

Private Sub Option4_Click()
Timer1.Interval = 1
End Sub

Private Sub Timer1_Timer()
Dim z As Integer
For X = 0 To 249
For Y = 0 To 249
z = 0
If X > 0 Then
If a(X - 1, Y) = True Then z = z + 1
If Y > 0 Then
If a(X - 1, Y - 1) = True Then z = z + 1
End If
If Y < 249 Then
If a(X - 1, Y + 1) Then z = z + 1
End If
End If
If X < 249 Then
If a(X + 1, Y) = True Then z = z + 1
If Y > 0 Then
If a(X + 1, Y - 1) = True Then z = z + 1
End If
If Y < 249 Then
If a(X + 1, Y + 1) Then z = z + 1
End If
End If
If Y > 0 Then
If a(X, Y - 1) = True Then z = z + 1
End If
If Y < 249 Then
If a(X, Y + 1) Then z = z + 1
End If
If z = 2 Then
ElseIf z = 3 Then
b(X, Y) = True
Else
b(X, Y) = False
End If
Next Y
Next X
For X = 0 To 249
For Y = 0 To 249
a(X, Y) = b(X, Y)
Next Y
Next X
Call DRAWING
End Sub
