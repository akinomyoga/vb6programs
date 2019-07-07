VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "この曲の設定"
   ClientHeight    =   2535
   ClientLeft      =   4815
   ClientTop       =   1845
   ClientWidth     =   2610
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "適用"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "音符"
      TabPicture(0)   =   "音楽1設定.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "速さ"
      TabPicture(1)   =   "音楽1設定.frx":0112
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "曲情報"
      TabPicture(2)   =   "音楽1設定.frx":0224
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text2"
      Tab(2).Control(1)=   "Text1"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label1"
      Tab(2).ControlCount=   4
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "音楽1設定.frx":0240
         Left            =   -73920
         List            =   "音楽1設定.frx":024A
         TabIndex        =   13
         Text            =   "Andante"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   -74400
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   -74400
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "三連符を使う"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "最小音符(一升の長さ)"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "十六分"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "八分"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "四分"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label2 
         Caption         =   "作者"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "曲名"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa(1) As Boolean, bb(1) As Integer

Public Sub shokika()
Option1(Form1.omp_mode1 Mod 3).Value = True
If Int(Form1.omp_mode1 / 3) Then Check1.Value = True Else Check1.Value = False
End Sub

Public Sub setting()
If aa(0) = True Then
aa(0) = False
Select Case bb(0)
Case 0: set0 4, 3, 2, 0, 1, 0, 0, 0
Case 1: set0 8, 6, 4, 3, 2, 0, 1, 0
Case 2: set0 16, 12, 8, 6, 4, 3, 2, 1
Case 3: set0 12, 9, 6, 0, 3, 0, 0, 0
Case 4: set0 24, 18, 12, 9, 6, 0, 3, 0
Case 5: set0 48, 36, 24, 18, 12, 9, 6, 3
End Select
Form1.omp_mode1 = bb(0)
End If
End Sub

Public Sub set0(a16, a12, a8, a6, a4, a3, a2, a1)
Form1.set_omp_long 1, (a16)
Form1.set_omp_long 2, (a8)
Form1.set_omp_long 3, (a12)
Form1.set_omp_long 4, (a4)
Form1.set_omp_long 5, (a6)
If a6 = 0 Then Form1.mn_omps(5).Enabled = False Else Form1.mn_omps(5).Enabled = True
Form1.set_omp_long 6, (a2)
If a2 = 0 Then Form1.mn_omps(6).Enabled = False Else Form1.mn_omps(6).Enabled = True
Form1.set_omp_long 7, (a3)
If a3 = 0 Then Form1.mn_omps(7).Enabled = False Else Form1.mn_omps(7).Enabled = True
Form1.set_omp_long 8, (a1)
If a1 = 0 Then Form1.mn_omps(8).Enabled = False Else Form1.mn_omps(8).Enabled = True
End Sub

Private Sub Command1_Click()
Call setting
Form1.Enabled = True
Form3.Hide
Form1.omp_mode = Form1.omp_mode
End Sub

Private Sub Command2_Click()
Call setting
Form1.omp_mode = Form1.omp_mode
End Sub

Private Sub Command3_Click()
Form1.Enabled = True
Form3.Hide
Call shokika
End Sub

Private Sub Form_Load()
Call shokika
End Sub

Private Sub Option1_Click(Index As Integer)
aa(0) = True
If Check1.Value = True Then X = 3
If Option1(Index).Value = True Then bb(0) = Index + X
End Sub

