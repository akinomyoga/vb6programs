VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   405
   ClientTop       =   2070
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8460
   Begin VB.CommandButton Command13 
      Caption         =   "病薬"
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "餌"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "持ち物"
      Height          =   5175
      Left            =   5520
      TabIndex        =   11
      Top             =   1200
      Width           =   2415
      Begin VB.Label Label8 
         Caption         =   "魚"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "肥料"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "農薬"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "種"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "栄養"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "牛乳"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "大牛"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "子牛"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "子供"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "収穫"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "収穫"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "雑草"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "掃除"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "肥料"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "栄養"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "農薬"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "病薬"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "水"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "餌"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   15
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '塗りつぶし
      Height          =   1215
      Left            =   120
      Shape           =   2  '楕円
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   49
      Left            =   4920
      Picture         =   "牧場.frx":0000
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   48
      Left            =   4920
      Picture         =   "牧場.frx":0C42
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   47
      Left            =   4920
      Picture         =   "牧場.frx":1884
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   46
      Left            =   4920
      Picture         =   "牧場.frx":24C6
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   45
      Left            =   4920
      Picture         =   "牧場.frx":3108
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   44
      Left            =   4440
      Picture         =   "牧場.frx":3D4A
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   43
      Left            =   4440
      Picture         =   "牧場.frx":498C
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   42
      Left            =   4440
      Picture         =   "牧場.frx":55CE
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   41
      Left            =   4440
      Picture         =   "牧場.frx":6210
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   40
      Left            =   4440
      Picture         =   "牧場.frx":6E52
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   39
      Left            =   3960
      Picture         =   "牧場.frx":7A94
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   38
      Left            =   3960
      Picture         =   "牧場.frx":86D6
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   37
      Left            =   3960
      Picture         =   "牧場.frx":9318
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   36
      Left            =   3960
      Picture         =   "牧場.frx":9F5A
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   35
      Left            =   3960
      Picture         =   "牧場.frx":AB9C
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   34
      Left            =   3480
      Picture         =   "牧場.frx":B7DE
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   33
      Left            =   3480
      Picture         =   "牧場.frx":C420
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   32
      Left            =   3480
      Picture         =   "牧場.frx":D062
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   31
      Left            =   3480
      Picture         =   "牧場.frx":DCA4
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   30
      Left            =   3480
      Picture         =   "牧場.frx":E8E6
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   29
      Left            =   3000
      Picture         =   "牧場.frx":F528
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   28
      Left            =   3000
      Picture         =   "牧場.frx":1016A
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   27
      Left            =   3000
      Picture         =   "牧場.frx":10DAC
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   26
      Left            =   3000
      Picture         =   "牧場.frx":119EE
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   25
      Left            =   3000
      Picture         =   "牧場.frx":12630
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   24
      Left            =   4920
      Picture         =   "牧場.frx":13272
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   23
      Left            =   4920
      Picture         =   "牧場.frx":13EB4
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   22
      Left            =   4920
      Picture         =   "牧場.frx":14AF6
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   21
      Left            =   4920
      Picture         =   "牧場.frx":15738
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   20
      Left            =   4920
      Picture         =   "牧場.frx":1637A
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   19
      Left            =   4440
      Picture         =   "牧場.frx":16FBC
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   18
      Left            =   4440
      Picture         =   "牧場.frx":17BFE
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   17
      Left            =   4440
      Picture         =   "牧場.frx":18840
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   16
      Left            =   4440
      Picture         =   "牧場.frx":19482
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   15
      Left            =   4440
      Picture         =   "牧場.frx":1A0C4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   14
      Left            =   3960
      Picture         =   "牧場.frx":1AD06
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   13
      Left            =   3960
      Picture         =   "牧場.frx":1B948
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   3960
      Picture         =   "牧場.frx":1C58A
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   11
      Left            =   3960
      Picture         =   "牧場.frx":1D1CC
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   10
      Left            =   3960
      Picture         =   "牧場.frx":1DE0E
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   9
      Left            =   3480
      Picture         =   "牧場.frx":1EA50
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   3480
      Picture         =   "牧場.frx":1F692
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   3480
      Picture         =   "牧場.frx":202D4
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   3480
      Picture         =   "牧場.frx":20F16
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   3480
      Picture         =   "牧場.frx":21B58
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   3000
      Picture         =   "牧場.frx":2279A
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   3000
      Picture         =   "牧場.frx":233DC
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3000
      Picture         =   "牧場.frx":2401E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   3000
      Picture         =   "牧場.frx":24C60
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   3000
      Picture         =   "牧場.frx":258A2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   39
      Left            =   2040
      Picture         =   "牧場.frx":260E4
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   38
      Left            =   2040
      Picture         =   "牧場.frx":26D26
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   37
      Left            =   2040
      Picture         =   "牧場.frx":27968
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   36
      Left            =   2040
      Picture         =   "牧場.frx":285AA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   35
      Left            =   2040
      Picture         =   "牧場.frx":291EC
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   34
      Left            =   2040
      Picture         =   "牧場.frx":29E2E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   33
      Left            =   2040
      Picture         =   "牧場.frx":2AA70
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   32
      Left            =   2040
      Picture         =   "牧場.frx":2B6B2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   31
      Left            =   1560
      Picture         =   "牧場.frx":2C2F4
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   30
      Left            =   1560
      Picture         =   "牧場.frx":2CF36
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   29
      Left            =   1560
      Picture         =   "牧場.frx":2DB78
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   28
      Left            =   1560
      Picture         =   "牧場.frx":2E7BA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   27
      Left            =   1560
      Picture         =   "牧場.frx":2F3FC
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   26
      Left            =   1560
      Picture         =   "牧場.frx":3003E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   25
      Left            =   1560
      Picture         =   "牧場.frx":30C80
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   24
      Left            =   1560
      Picture         =   "牧場.frx":318C2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   23
      Left            =   1080
      Picture         =   "牧場.frx":32504
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   22
      Left            =   1080
      Picture         =   "牧場.frx":33146
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   21
      Left            =   1080
      Picture         =   "牧場.frx":33D88
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   20
      Left            =   1080
      Picture         =   "牧場.frx":349CA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   19
      Left            =   1080
      Picture         =   "牧場.frx":3560C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   18
      Left            =   1080
      Picture         =   "牧場.frx":3624E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   17
      Left            =   1080
      Picture         =   "牧場.frx":36E90
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   16
      Left            =   1080
      Picture         =   "牧場.frx":37AD2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   15
      Left            =   600
      Picture         =   "牧場.frx":38714
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   14
      Left            =   600
      Picture         =   "牧場.frx":39356
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   600
      Picture         =   "牧場.frx":39F98
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   600
      Picture         =   "牧場.frx":3ABDA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   600
      Picture         =   "牧場.frx":3B81C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   600
      Picture         =   "牧場.frx":3C45E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   600
      Picture         =   "牧場.frx":3D0A0
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   600
      Picture         =   "牧場.frx":3DCE2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   120
      Picture         =   "牧場.frx":3E924
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   120
      Picture         =   "牧場.frx":3F566
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   120
      Picture         =   "牧場.frx":401A8
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   120
      Picture         =   "牧場.frx":40DEA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "牧場.frx":41A2C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "牧場.frx":4266E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "牧場.frx":432B0
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "牧場.frx":43EF2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Menu MNfile 
      Caption         =   "ファイル"
      Begin VB.Menu MNclear 
         Caption         =   "最初から"
         Shortcut        =   ^N
      End
      Begin VB.Menu MNopen 
         Caption         =   "開く"
         Shortcut        =   ^O
      End
      Begin VB.Menu MNsena 
         Caption         =   "-"
      End
      Begin VB.Menu MNhozon 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu MNsinfile 
         Caption         =   "名前を付けて保存"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNsenb 
         Caption         =   "-"
      End
      Begin VB.Menu MNend 
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
Private Sub MNend_click()
End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub MNhozon_Click()
Form2.Show
End Sub

