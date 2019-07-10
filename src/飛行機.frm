VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "飛行機"
   ClientHeight    =   5265
   ClientLeft      =   3180
   ClientTop       =   3105
   ClientWidth     =   7110
   Icon            =   "飛行機.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7110
   Begin VB.Timer Timer3 
      Left            =   1800
      Top             =   4800
   End
   Begin VB.Timer Timer2 
      Left            =   1320
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   4800
   End
   Begin KBasic.SpinButton SpinButton1 
      Height          =   5275
      Left            =   0
      Top             =   0
      Width           =   255
      _extentx        =   0
      _extenty        =   0
      min             =   4920
      max             =   120
      smallchange     =   120
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   255
      Picture         =   "飛行機.frx":27A2
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "終わり"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   72
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   19
      Left            =   6240
      Picture         =   "飛行機.frx":2F64
      ToolTipText     =   "敵だー!!!"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   18
      Left            =   6240
      Picture         =   "飛行機.frx":3AA6
      ToolTipText     =   "敵だー!!!"
      Top             =   4560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   17
      Left            =   5760
      Picture         =   "飛行機.frx":45E8
      ToolTipText     =   "敵だー!!!"
      Top             =   4560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   16
      Left            =   5760
      Picture         =   "飛行機.frx":512A
      ToolTipText     =   "敵だー!!!"
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   15
      Left            =   5760
      Picture         =   "飛行機.frx":5C6C
      ToolTipText     =   "敵だー!!!"
      Top             =   720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   14
      Left            =   5760
      Picture         =   "飛行機.frx":67AE
      ToolTipText     =   "敵だー!!!"
      Top             =   1200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   13
      Left            =   5760
      Picture         =   "飛行機.frx":72F0
      ToolTipText     =   "敵だー!!!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "飛行機.frx":7E32
      ToolTipText     =   "敵だー!!!"
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   11
      Left            =   5760
      Picture         =   "飛行機.frx":8974
      ToolTipText     =   "敵だー!!!"
      Top             =   2640
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   10
      Left            =   5760
      Picture         =   "飛行機.frx":94B6
      ToolTipText     =   "敵だー!!!"
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   9
      Left            =   5760
      Picture         =   "飛行機.frx":9FF8
      ToolTipText     =   "敵だー!!!"
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   5760
      Picture         =   "飛行機.frx":AB3A
      ToolTipText     =   "敵だー!!!"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   6240
      Picture         =   "飛行機.frx":B67C
      ToolTipText     =   "敵だー!!!"
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   6240
      Picture         =   "飛行機.frx":C1BE
      ToolTipText     =   "敵だー!!!"
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "飛行機.frx":CD00
      ToolTipText     =   "敵だー!!!"
      Top             =   2640
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   6240
      Picture         =   "飛行機.frx":D842
      ToolTipText     =   "敵だー!!!"
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   6240
      Picture         =   "飛行機.frx":E384
      ToolTipText     =   "敵だー!!!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   6240
      Picture         =   "飛行機.frx":EEC6
      ToolTipText     =   "敵だー!!!"
      Top             =   1200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   6240
      Picture         =   "飛行機.frx":FA08
      ToolTipText     =   "敵だー!!!"
      Top             =   720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   6240
      Picture         =   "飛行機.frx":1054A
      ToolTipText     =   "敵だー!!!"
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   720
      X2              =   7200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "点"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   4800
      Width           =   375
   End
   Begin VB.Menu file 
      Caption         =   "ファイル"
      Begin VB.Menu mnstart 
         Caption         =   "始めから"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnend 
         Caption         =   "終了"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnmouse 
      Caption         =   "マウス"
      Begin VB.Menu mnsou 
         Caption         =   "操縦"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu am 
      Caption         =   "ヘルプ"
      Begin VB.Menu bm 
         Caption         =   "ファイルの始めからについて・・・一番最初からやりなおす"
         Enabled         =   0   'False
      End
      Begin VB.Menu cm 
         Caption         =   "-"
      End
      Begin VB.Menu d 
         Caption         =   "ファイルの終了について・・・このゲームをやめたい時にお"
         Enabled         =   0   'False
      End
      Begin VB.Menu e 
         Caption         =   "すと、このプログラムが終了する"
         Enabled         =   0   'False
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu g 
         Caption         =   "マウスの操縦について・・・チェックを入れるとマウスで自分"
         Enabled         =   0   'False
      End
      Begin VB.Menu h 
         Caption         =   "を動かせるが、簡単。チェックをはずすと、キーボードの矢"
         Enabled         =   0   'False
      End
      Begin VB.Menu i 
         Caption         =   "印キーでしか自分を動かすことが出来ないので、難しい。"
         Enabled         =   0   'False
      End
      Begin VB.Menu t 
         Caption         =   "-"
      End
      Begin VB.Menu j 
         Caption         =   "-"
      End
      Begin VB.Menu k 
         Caption         =   "ゲームのやりかたについて・・・"
         Enabled         =   0   'False
      End
      Begin VB.Menu l 
         Caption         =   "１．ファイルの始めからを選ぶ"
         Enabled         =   0   'False
      End
      Begin VB.Menu m 
         Caption         =   "２．マウス又はキーボードの矢印キーで上下に移動する。"
         Enabled         =   0   'False
      End
      Begin VB.Menu n 
         Caption         =   "３．クリック（マウスの左のボタンを押して離す）すると敵を"
         Enabled         =   0   'False
      End
      Begin VB.Menu o 
         Caption         =   "　　攻撃出来る。"
         Enabled         =   0   'False
      End
      Begin VB.Menu p 
         Caption         =   "-"
      End
      Begin VB.Menu q 
         Caption         =   "得点の仕方について・・・"
         Enabled         =   0   'False
      End
      Begin VB.Menu r 
         Caption         =   "敵を倒すと１点もらえます。しかし、敵を逃してしまうと得点"
         Enabled         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "が一点減ってしまいます。"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim score As Integer
Dim x


Private Sub UTU()
Line1.Visible = True
Line1.Y1 = SpinButton1.Value + 165
Line1.Y2 = Line1.Y1
Timer1.Interval = 100
For c = 0 To 19
    If Image2(c).Visible = True Then
        If Line1.Y1 >= Image2(c).Top And Line1.Y1 <= Image2(c).Top + 480 Then
            Image2(c).Visible = False
            If Label2.Visible = False Then score = score + 1
            Label1.Caption = score
        End If
    End If
Next c
End Sub

Private Sub Form_Load()
b = 0
score = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then Call UTU
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Y >= 285 And Y <= 5085 Then
    If mnsou.Checked = True Then
        SpinButton1.Value = Y - 165
    End If
End If
End Sub

Private Sub sennsha()
Image2(b).Visible = True
Randomize
c = Int(Rnd * 37): If c = 37 Then c = 0
Image2(b).Top = c * 120 + 120
Image2(b).Left = 7200
End Sub

Private Sub kuria()
Label2.Visible = True
Timer2.Interval = 0
Timer3.Interval = 0
End Sub

Private Sub mnend_Click()
End
End Sub

Private Sub mnsou_Click()
If mnsou.Checked = True Then
mnsou.Checked = False
ElseIf mnsou.Checked = False Then
mnsou.Checked = True
End If
End Sub

Private Sub mnstart_Click()
Timer3.Interval = 200
Timer2.Interval = 3000
score = 0
Label2.Visible = False
Label1.Caption = 0
For c = 0 To 19
Image2(c).Visible = False
Next c
End Sub

Private Sub SpinButton1_Change()
Image1.Top = SpinButton1.Value
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Line1.Visible = False
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = Timer2.Interval - 20
If Timer2.Interval >= 1000 Then
    Call sennsha
ElseIf Timer2.Interval <= 800 Then
    Call kuria
End If
b = b + 1: If b = 20 Then b = 0
End Sub

Private Sub Timer3_Timer()
For c = 0 To 19
Image2(c).Left = Image2(c).Left - 240
If Image2(c).Left <= 120 Then
Image2(c).Left = 7200
If Image2(c).Visible = True Then score = score - 1
Image2(c).Visible = False
Label1.Caption = score
End If
Next c
End Sub
