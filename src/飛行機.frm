VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��s�@"
   ClientHeight    =   5265
   ClientLeft      =   3180
   ClientTop       =   3105
   ClientWidth     =   7110
   Icon            =   "��s�@.frx":0000
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
      Picture         =   "��s�@.frx":27A2
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�I���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Picture         =   "��s�@.frx":2F64
      ToolTipText     =   "�G���[!!!"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   18
      Left            =   6240
      Picture         =   "��s�@.frx":3AA6
      ToolTipText     =   "�G���[!!!"
      Top             =   4560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   17
      Left            =   5760
      Picture         =   "��s�@.frx":45E8
      ToolTipText     =   "�G���[!!!"
      Top             =   4560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   16
      Left            =   5760
      Picture         =   "��s�@.frx":512A
      ToolTipText     =   "�G���[!!!"
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   15
      Left            =   5760
      Picture         =   "��s�@.frx":5C6C
      ToolTipText     =   "�G���[!!!"
      Top             =   720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   14
      Left            =   5760
      Picture         =   "��s�@.frx":67AE
      ToolTipText     =   "�G���[!!!"
      Top             =   1200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   13
      Left            =   5760
      Picture         =   "��s�@.frx":72F0
      ToolTipText     =   "�G���[!!!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   5760
      Picture         =   "��s�@.frx":7E32
      ToolTipText     =   "�G���[!!!"
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   11
      Left            =   5760
      Picture         =   "��s�@.frx":8974
      ToolTipText     =   "�G���[!!!"
      Top             =   2640
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   10
      Left            =   5760
      Picture         =   "��s�@.frx":94B6
      ToolTipText     =   "�G���[!!!"
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   9
      Left            =   5760
      Picture         =   "��s�@.frx":9FF8
      ToolTipText     =   "�G���[!!!"
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   5760
      Picture         =   "��s�@.frx":AB3A
      ToolTipText     =   "�G���[!!!"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   6240
      Picture         =   "��s�@.frx":B67C
      ToolTipText     =   "�G���[!!!"
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   6240
      Picture         =   "��s�@.frx":C1BE
      ToolTipText     =   "�G���[!!!"
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "��s�@.frx":CD00
      ToolTipText     =   "�G���[!!!"
      Top             =   2640
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   6240
      Picture         =   "��s�@.frx":D842
      ToolTipText     =   "�G���[!!!"
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   6240
      Picture         =   "��s�@.frx":E384
      ToolTipText     =   "�G���[!!!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   6240
      Picture         =   "��s�@.frx":EEC6
      ToolTipText     =   "�G���[!!!"
      Top             =   1200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   6240
      Picture         =   "��s�@.frx":FA08
      ToolTipText     =   "�G���[!!!"
      Top             =   720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   6240
      Picture         =   "��s�@.frx":1054A
      ToolTipText     =   "�G���[!!!"
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�_"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�t�@�C��"
      Begin VB.Menu mnstart 
         Caption         =   "�n�߂���"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnend 
         Caption         =   "�I��"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnmouse 
      Caption         =   "�}�E�X"
      Begin VB.Menu mnsou 
         Caption         =   "���c"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu am 
      Caption         =   "�w���v"
      Begin VB.Menu bm 
         Caption         =   "�t�@�C���̎n�߂���ɂ��āE�E�E��ԍŏ�������Ȃ���"
         Enabled         =   0   'False
      End
      Begin VB.Menu cm 
         Caption         =   "-"
      End
      Begin VB.Menu d 
         Caption         =   "�t�@�C���̏I���ɂ��āE�E�E���̃Q�[������߂������ɂ�"
         Enabled         =   0   'False
      End
      Begin VB.Menu e 
         Caption         =   "���ƁA���̃v���O�������I������"
         Enabled         =   0   'False
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu g 
         Caption         =   "�}�E�X�̑��c�ɂ��āE�E�E�`�F�b�N������ƃ}�E�X�Ŏ���"
         Enabled         =   0   'False
      End
      Begin VB.Menu h 
         Caption         =   "�𓮂����邪�A�ȒP�B�`�F�b�N���͂����ƁA�L�[�{�[�h�̖�"
         Enabled         =   0   'False
      End
      Begin VB.Menu i 
         Caption         =   "��L�[�ł��������𓮂������Ƃ��o���Ȃ��̂ŁA����B"
         Enabled         =   0   'False
      End
      Begin VB.Menu t 
         Caption         =   "-"
      End
      Begin VB.Menu j 
         Caption         =   "-"
      End
      Begin VB.Menu k 
         Caption         =   "�Q�[���̂�肩���ɂ��āE�E�E"
         Enabled         =   0   'False
      End
      Begin VB.Menu l 
         Caption         =   "�P�D�t�@�C���̎n�߂����I��"
         Enabled         =   0   'False
      End
      Begin VB.Menu m 
         Caption         =   "�Q�D�}�E�X���̓L�[�{�[�h�̖��L�[�ŏ㉺�Ɉړ�����B"
         Enabled         =   0   'False
      End
      Begin VB.Menu n 
         Caption         =   "�R�D�N���b�N�i�}�E�X�̍��̃{�^���������ė����j����ƓG��"
         Enabled         =   0   'False
      End
      Begin VB.Menu o 
         Caption         =   "�@�@�U���o����B"
         Enabled         =   0   'False
      End
      Begin VB.Menu p 
         Caption         =   "-"
      End
      Begin VB.Menu q 
         Caption         =   "���_�̎d���ɂ��āE�E�E"
         Enabled         =   0   'False
      End
      Begin VB.Menu r 
         Caption         =   "�G��|���ƂP�_���炦�܂��B�������A�G�𓦂��Ă��܂��Ɠ��_"
         Enabled         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "����_�����Ă��܂��܂��B"
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
