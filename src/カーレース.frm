VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   3795
   ClientTop       =   3330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6750
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3120
      Picture         =   "カーレース.frx":0000
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   6735
   End
   Begin VB.Menu file 
      Caption         =   "ファイル"
      Begin VB.Menu start 
         Caption         =   "スタート"
         Shortcut        =   ^S
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu stop 
         Caption         =   "ストップ"
         Shortcut        =   ^D
      End
      Begin VB.Menu restart 
         Caption         =   "再開"
         Shortcut        =   ^F
      End
      Begin VB.Menu ddd 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "終了"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(29), d(29), n(59) As Integer, nc As Integer
Private Sub end_Click()
MsgBox "終了します。"
End
End Sub

Private Sub Form_Load()
n(0) = 0
n(1) = 0
n(2) = 0
n(3) = 0
n(4) = 0
n(5) = 0
n(6) = 0
n(7) = 0
n(8) = 0
n(9) = 0
n(10) = 0
n(11) = 0
n(12) = 0
n(13) = 0
n(14) = 0
n(15) = 0
n(16) = 0
n(17) = 1
n(18) = 3
n(19) = 5
n(20) = 7
n(21) = 9
n(22) = 11
n(23) = 13
n(24) = 15
n(25) = 17
n(26) = 19
n(27) = 15
n(28) = 11
n(29) = 7
n(30) = 3
n(40) = 5
n(41) = -5
n(42) = 10
n(43) = -10
n(44) = 15
n(45) = -15
n(46) = 10
n(48) = 5


End Sub

Private Sub restart_Click()
Timer1.Interval = 200
End Sub

Private Sub start_Click()
Timer1.Interval = 10
nc = 0
End Sub

Private Sub stop_Click()
Timer1.Interval = 0
End Sub

Private Sub Timer1_Timer()
On Error GoTo err
For b = 29 To -6 Step -1
dc = Int((29 - b) / 5)
nd = Int(nc / 5)
bc = b + nc Mod 5
If bc < 30 And bc >= 0 Then d(bc) = n(nd + dc + 1) - (n(nd + dc + 1) - n(nd + dc)) / 5 * ((b + 10) Mod 5)
Next b
nc = nc + 1
Label1.Caption = "あと" & 266 - nc & "cm"
Cls
DrawWidth = 8
l = 2760
r = 3960
For b = 0 To 29
e = Sin((90 - b * 3) * 3.14 / 180) * d(b) * 150
l = l - 90
r = r + 90
m = 2520 + b * 80
Line (l + e, m)-(r + e, m), RGB(150, 150, 150)
PSet (l + e, m), RGB(0, 0, 0)
PSet (r + e, m), RGB(0, 0, 0)
Next b
Image1.Top = Image1.Top
Exit Sub
err:
MsgBox "ゴール！"
nc = 0
End Sub
