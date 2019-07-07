VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   3930
   ClientTop       =   3300
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7605
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.HScrollBar HScroll1 
      DragMode        =   1  '自動
      Height          =   255
      Left            =   1080
      Max             =   600
      TabIndex        =   0
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7080
      Top             =   720
   End
   Begin MSForms.ScrollBar ScrollBar1 
      DragMode        =   1  '自動
      Height          =   2055
      Left            =   6480
      TabIndex        =   4
      Top             =   3000
      Width           =   495
      Size            =   "873;3625"
      Min             =   10
      Max             =   -10
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   2055
      Left            =   6960
      TabIndex        =   5
      Top             =   3000
      Width           =   495
      Size            =   "873;3625"
      Min             =   -10
      Max             =   10
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   120
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7560
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   1200
      Picture         =   "電車.frx":0000
      Top             =   1800
      Width           =   4995
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "電車.frx":3622
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   7575
   End
   Begin VB.Menu file 
      Caption         =   "ファイル"
      Begin VB.Menu start 
         Caption         =   "最初から"
         Shortcut        =   ^N
      End
      Begin VB.Menu bou1 
         Caption         =   "-"
      End
      Begin VB.Menu read 
         Caption         =   "読み込み"
         Shortcut        =   ^O
      End
      Begin VB.Menu bou2 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "終了"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Single

Private Sub end_Click()
End
End Sub

Private Sub Form_Load()
a = 0: b = 100000
End Sub

Private Sub SpinButton1_Change()
ScrollBar1.Value = SpinButton1.Value
End Sub

Private Sub start_Click()
b = 100000
HScroll1.Value = 0
ScrollBar1.Value = 0
SpinButton1.Value = 0
HScroll1.Max = 600
Label1.ForeColor = RGB(0, 0, 255)
Label1.BackColor = RGB(0, 255, 255)
End Sub

Private Sub Timer1_Timer()
If HScroll1.Value - -1 * ScrollBar1.Value >= 0 And HScroll1.Value - -1 * ScrollBar1.Value <= HScroll1.Max Then
HScroll1.Value = HScroll1.Value - -1 * ScrollBar1.Value
End If
Label3.Caption = ScrollBar1.Value
If Label3.Caption > 0 Then
Label3.ForeColor = RGB(0, 255, 0)
ElseIf Label3.Caption < 0 Then
Label3.ForeColor = RGB(255, 0, 0)
Else
Label3.ForeColor = RGB(0, 0, 255)
End If
Label4.Caption = "速度　" & 5 * HScroll1.Value & "cm/秒：" & Int(9 / 50 * HScroll1.Value) & "km/時"

If a = 0 Then
Image1.Left = 1200 - HScroll1.Value
a = 1
Else
Image1.Left = 1200
a = 0
End If

b = b - HScroll1.Value
If b > 0 Then
Label1.Caption = "次の駅まであと" & b & "cm"
Else
    If b > -6 Then
    Label1.Caption = "合格範囲！//" & -1 * b & "cm オーバー"
    Label1.ForeColor = RGB(0, 127, 0)
        If HScroll1.Value = 0 Then
        Label1.Caption = "クリアーおめでとう！　おわり"
        HScroll1.Max = 0
        End If
    Else
    Label1.ForeColor = RGB(255, 0, 0)
    Label1.BackColor = RGB(255, 255, 0)
    Label1.Caption = "「まってー電車さん！乗れないよー！」//" & -1 * b & "cm オーバー"
        If HScroll1.Value = 0 Then
        Label1.Caption = "お客さんが・・・怒っている！　おわり"
        HScroll1.Max = 0
        End If
    End If
End If
Label2.Width = (100000 - b) / 100000 * 3255

End Sub
