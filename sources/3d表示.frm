VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "三点透視"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.VScrollBar VScroll3 
      Height          =   7575
      LargeChange     =   480
      Left            =   9840
      Max             =   -30000
      Min             =   30000
      SmallChange     =   240
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   540
      Left            =   4200
      Max             =   -30000
      Min             =   30000
      SmallChange     =   240
      TabIndex        =   14
      Top             =   7680
      Width           =   5655
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3495
      LargeChange     =   50
      Left            =   9240
      Max             =   5000
      Min             =   50
      SmallChange     =   50
      TabIndex        =   11
      Top             =   1560
      Value           =   500
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      LargeChange     =   2
      Left            =   8760
      Max             =   2000
      Min             =   100
      SmallChange     =   2
      TabIndex        =   10
      Top             =   1560
      Value           =   100
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   4920
      ScaleHeight     =   2235
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   5280
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   5
         ItemData        =   "3d表示.frx":0000
         Left            =   3600
         List            =   "3d表示.frx":0002
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   4
         ItemData        =   "3d表示.frx":0004
         Left            =   2880
         List            =   "3d表示.frx":0006
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   3
         ItemData        =   "3d表示.frx":0008
         Left            =   2160
         List            =   "3d表示.frx":000A
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   2
         ItemData        =   "3d表示.frx":000C
         Left            =   1440
         List            =   "3d表示.frx":000E
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   1
         ItemData        =   "3d表示.frx":0010
         Left            =   720
         List            =   "3d表示.frx":0012
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   1680
         Index           =   0
         ItemData        =   "3d表示.frx":0014
         Left            =   0
         List            =   "3d表示.frx":0016
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "始点　　　　　　　　　　終点                      X        Y        Z        X       Y        Z"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "表示"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "入力..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "large   far"
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "small  near"
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1(3, 1) As Integer, x1 As Double, y1 As Double

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form1.Cls
If Form2.p4count <> 0 Then
For a = 0 To Form2.p4count - 1
Call Get2D(List1(0).List(a), List1(1).List(a), List1(2).List(a))
xa = x1
ya = y1
Call Get2D(List1(3).List(a), List1(4).List(a), List1(5).List(a))
xb = x1
yb = y1
Form1.Line (xa, ya)-(xb, yb), RGB(0, 0, 0)
Next a
End If
End Sub

Private Sub Form_Load()
p1(1, 0) = VScroll2.Value * 2
p1(1, 1) = VScroll2.Value * 1
p1(2, 0) = -VScroll2.Value * 2
p1(2, 1) = VScroll2.Value * 1
p1(3, 0) = 0
p1(3, 1) = VScroll2.Value * 5 ^ (1 / 2)
End Sub

Private Sub Form_Resize()
p1(0, 0) = Form1.Width / 2
p1(0, 1) = Form1.Height / 10
End Sub

Public Sub Get2D(x As Double, y As Double, z As Double)
xa = (1 - (995 / 1000) ^ (x - 1500))
ya = (1 - (995 / 1000) ^ (y - 1500))
za = (1 - (995 / 1000) ^ (z - 1500))
x1 = (p1(1, 0) * xa + p1(2, 0) * ya + p1(3, 0) * za) * VScroll1.Value * 0.5 / VScroll2.Value ^ 1.3 + p1(0, 0) + HScroll1.Value
y1 = (p1(1, 1) * xa + p1(2, 1) * ya + p1(3, 1) * za) * VScroll1.Value * 0.5 / VScroll2.Value ^ 1.3 + p1(0, 1) + VScroll3.Value
End Sub

Private Sub HScroll1_Change()
Call Command2_Click
End Sub

Private Sub VScroll1_Change()
Call Command2_Click
End Sub

Private Sub VScroll2_Change()
Call Form_Load
Call Command2_Click
End Sub

Private Sub VScroll3_Change()
Call Command2_Click
End Sub
