VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "町"
   ClientHeight    =   6930
   ClientLeft      =   4125
   ClientTop       =   2655
   ClientWidth     =   8145
   Icon            =   "町.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8145
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":27A2
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   36
      Top             =   4680
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   4800
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   1080
      Max             =   5
      TabIndex        =   27
      Top             =   360
      Width           =   3375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      Left            =   720
      Max             =   5
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "終了"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":2FE4
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   4
      ToolTipText     =   "<\10000万 ﾏﾝｼｮﾝ>  4万人  人気2"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":3826
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   3
      ToolTipText     =   "<\1000万　工場>  100万円"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":4068
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      ToolTipText     =   "<\3000万 施設>  人気3  10万円"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":48AA
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      ToolTipText     =   "<\500万 道路>  人気5"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   4440
      Picture         =   "町.frx":50EC
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      ToolTipText     =   "<\5000万　ﾏﾝｼｮﾝ>  2万人  人気1"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   42
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "日数"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   40
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   39
      ToolTipText     =   "人口の最高の合計"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   38
      ToolTipText     =   "人口の最高"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   37
      ToolTipText     =   "人口の最高"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "100000000"
      Height          =   255
      Left            =   2280
      TabIndex        =   35
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "予算"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   34
      Top             =   5280
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      Height          =   3735
      Left            =   7080
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "人気"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   33
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   32
      ToolTipText     =   "町の人気"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   31
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   30
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   29
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   4095
      Left            =   5160
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   26
      ToolTipText     =   "建物の数"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   25
      ToolTipText     =   "この建物の数"
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   24
      ToolTipText     =   "この建物の数"
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   23
      ToolTipText     =   "この建物の数"
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   22
      ToolTipText     =   "この建物の数"
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   21
      ToolTipText     =   "この建物の数"
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "数"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF00FF&
      Height          =   3735
      Left            =   6120
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   19
      ToolTipText     =   "合計"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   18
      ToolTipText     =   "合計"
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   17
      ToolTipText     =   "合計"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   16
      ToolTipText     =   "合計"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   15
      ToolTipText     =   "合計"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "合計"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   13
      ToolTipText     =   "総人口"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   12
      ToolTipText     =   "一つあたり"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "一つあたり"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   10
      ToolTipText     =   "一つあたり"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "一つあたり"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "一つあたり"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "人口"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   1080
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   1560
      X2              =   1560
      Y1              =   720
      Y2              =   4080
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   35
      Left            =   3720
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   34
      Left            =   3720
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   33
      Left            =   3720
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   32
      Left            =   3720
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   31
      Left            =   3720
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   30
      Left            =   3720
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   29
      Left            =   3240
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   28
      Left            =   3240
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   27
      Left            =   3240
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   26
      Left            =   3240
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   25
      Left            =   3240
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   24
      Left            =   3240
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   23
      Left            =   2760
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   22
      Left            =   2760
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   21
      Left            =   2760
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   20
      Left            =   2760
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   19
      Left            =   2760
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   18
      Left            =   2760
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   17
      Left            =   2280
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   16
      Left            =   2280
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   15
      Left            =   2280
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   14
      Left            =   2280
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   13
      Left            =   2280
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   12
      Left            =   2280
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   11
      Left            =   1800
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   10
      Left            =   1800
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   9
      Left            =   1800
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   8
      Left            =   1800
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   1800
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   1800
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   1320
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   1320
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   1320
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   1320
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   1320
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   1320
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imgd(0 To 5, 0 To 5)
Dim a
Dim b
Dim ka As Integer
Dim kb As Integer
Dim kd As Integer
Dim kdd As Integer
Dim ke As Integer
Private Sub Command1_Click()
If Label5.Caption >= 50000000 Then
Call Herutatemono
Image1(a + 6 * b).Picture = Command1.Picture
imgd(a, b) = 1
Label4(0).Caption = Label4(0).Caption + 1
Label4(5).Caption = Label4(5).Caption + 1
Label6(0).Caption = Label6(0).Caption + 1
Label5.Caption = Label5.Caption - 50000000
Call NinkiGoukei
Call MachiNoTeiin
End If
End Sub
Private Sub Command2_Click()
If Label5.Caption >= 5000000 Then
Call Herutatemono
Image1(a + 6 * b).Picture = Command2.Picture
imgd(a, b) = 2
Label4(1).Caption = Label4(1).Caption + 1
Label4(5).Caption = Label4(5).Caption + 1
Label6(1).Caption = Label6(1).Caption + 5
Label5.Caption = Label5.Caption - 5000000
Call NinkiGoukei
Call MachiNoTeiin
End If
End Sub
Private Sub Command3_Click()
If Label5.Caption >= 30000000 Then
Call Herutatemono
Image1(a + 6 * b).Picture = Command3.Picture
imgd(a, b) = 3
Label4(2).Caption = Label4(2).Caption + 1
Label4(5).Caption = Label4(5).Caption + 1
Label6(2).Caption = Label6(2).Caption + 3
Label5.Caption = Label5.Caption - 30000000
Call NinkiGoukei
Call MachiNoTeiin
End If
End Sub
Private Sub Command4_Click()
If Label5.Caption >= 10000000 Then
Call Herutatemono
Image1(a + 6 * b).Picture = Command4.Picture
imgd(a, b) = 4
Label4(3).Caption = Label4(3).Caption + 1
Label4(5).Caption = Label4(5).Caption + 1
Label5.Caption = Label5.Caption - 10000000
Call NinkiGoukei
Call MachiNoTeiin
End If
End Sub
Private Sub Command5_Click()
If Label5.Caption >= 100000000 Then
Call Herutatemono
Image1(a + 6 * b).Picture = Command5.Picture
imgd(a, b) = 5
Label4(4).Caption = Label4(4).Caption + 1
Label4(5).Caption = Label4(5).Caption + 1
Label6(4).Caption = Label6(4).Caption + 2
Label5.Caption = Label5.Caption - 100000000
Call NinkiGoukei
Call MachiNoTeiin
End If
End Sub
Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
ke = 0
End Sub

Private Sub Form_Load()
imgd(0, 0) = 0: imgd(0, 1) = 0: imgd(0, 2) = 0: imgd(0, 3) = 0: imgd(0, 4) = 0: imgd(0, 5) = 0
imgd(1, 0) = 0: imgd(1, 1) = 0: imgd(1, 2) = 0: imgd(1, 3) = 0: imgd(1, 4) = 0: imgd(1, 5) = 0
imgd(2, 0) = 0: imgd(2, 1) = 0: imgd(2, 2) = 0: imgd(2, 3) = 0: imgd(2, 4) = 0: imgd(2, 5) = 0
imgd(3, 0) = 0: imgd(3, 1) = 0: imgd(3, 2) = 0: imgd(3, 3) = 0: imgd(3, 4) = 0: imgd(3, 5) = 0
imgd(4, 0) = 0: imgd(4, 1) = 0: imgd(4, 2) = 0: imgd(4, 3) = 0: imgd(4, 4) = 0: imgd(4, 5) = 0
imgd(5, 0) = 0: imgd(5, 1) = 0: imgd(5, 2) = 0: imgd(5, 3) = 0: imgd(5, 4) = 0: imgd(5, 5) = 0
a = 0: b = 0
End Sub
Private Sub HScroll1_Change()
b = HScroll1.Value
Line1.X1 = 1560 + 480 * HScroll1.Value
Line1.X2 = 1560 + 480 * HScroll1.Value
End Sub
Private Sub Label2_Change(Index As Integer)
If Index = 5 Then
lc = 3 * Label4(0).Caption + Label4(1).Caption + 2 * Label4(2).Caption + Label4(3).Caption + 5 * Label4(4).Caption
If lc > 0 Then
Label2(0).Caption = Fix(3 * Label2(5).Caption / lc)
Label2(1).Caption = Fix(Label2(5).Caption / lc)
Label2(2).Caption = Fix(2 * Label2(5).Caption / lc)
Label2(3).Caption = Fix(Label2(5).Caption / lc)
Label2(4).Caption = Fix(5 * Label2(5).Caption / lc)
End If
Label3(0).Caption = Label4(0).Caption * Label2(0).Caption
Label3(1).Caption = Label4(1).Caption * Label2(1).Caption
Label3(2).Caption = Label4(2).Caption * Label2(2).Caption
Label3(3).Caption = Label4(3).Caption * Label2(3).Caption
Label3(4).Caption = Label4(4).Caption * Label2(4).Caption
End If
End Sub
Private Sub MachiNoTeiin() '町に住める人の最高の人数を出し直す。
Dim ma As Long: ma = 20000 * Label4(0).Caption: Label7(0).Caption = ma
Dim mb As Long: mb = 40000 * Label4(4).Caption: Label7(4).Caption = mb
Label7(5).Caption = ma + mb
End Sub
Private Sub NinkiGoukei() '人気の合計を計算し直す。
Dim na As Integer: Dim nb As Integer: Dim nc As Integer: Dim nd As Integer: Dim ne As Integer
na = Label6(0).Caption: nb = Label6(1).Caption: nc = Label6(2).Caption: nd = Label6(3).Caption: ne = Label6(4).Caption
Label6(5).Caption = na + nb + nc + nd + ne
End Sub

Private Sub Herutatemono() '建物が減った時の計算です。
If imgd(a, b) <= 5 And imgd(a, b) >= 1 Then
Label4(5).Caption = Label4(5).Caption - 1
If imgd(a, b) = 1 Then
Label4(0).Caption = Label4(0).Caption - 1
Label6(0).Caption = Label6(0).Caption - 1
ElseIf imgd(a, b) = 2 Then
Label4(1).Caption = Label4(1).Caption - 1
Label6(1).Caption = Label6(1).Caption - 5
ElseIf imgd(a, b) = 3 Then
Label4(2).Caption = Label4(2).Caption - 1
Label6(2).Caption = Label6(2).Caption - 3
ElseIf imgd(a, b) = 4 Then
Label4(3).Caption = Label4(3).Caption - 1
ElseIf imgd(a, b) = 5 Then
Label4(4).Caption = Label4(4).Caption - 1
Label6(4).Caption = Label6(4).Caption - 2
End If
End If
End Sub

Private Sub Kaji()
If kd >= 0 And kd <= 35 Then
    Image1(kd).Picture = Command7.Picture
    kc = imgd(Fix(kd / 6), kd Mod 6)
    If kc > 0 Then
        Label4(kc - 1).Caption = Label4(kc - 1).Caption - 1
        Label4(5).Caption = Label4(5).Caption - 1
        Label2(5).Caption = Label2(5).Caption - Label2(kc - 1).Caption
        Label6(0).Caption = Label4(0).Caption: Label6(1).Caption = Label4(1).Caption * 5
        Label6(2).Caption = Label4(2).Caption * 3: Label6(4).Caption = Label4(4).Caption * 2
        Dim w As Integer: w = Label6(0).Caption: Dim x As Integer: x = Label6(1).Caption
        Dim y As Integer: y = Label6(2).Caption: Dim z As Integer: z = Label6(4).Caption
        Label6(5).Caption = w + x + y + z
        imgd(Fix(kd / 6), kd Mod 6) = 0
        Call MachiNoTeiin
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim ta As Long: ta = Label8.Caption
Dim tb As Long: tb = 1
Label8.Caption = ta + tb
'人口の増加
Dim h As Long: Dim i As Currency
h = Label2(5).Caption: i = Label6(5).Caption
h = h + i: h = h * (1000 + i) / 1000
If Int(h) <= Label7(5).Caption Then
Label2(5).Caption = Int(h)
Else
Label2(5).Caption = Label7(5).Caption
End If
'儲け
Dim j As Integer: j = Label4(2).Caption
Dim k As Integer: k = Label4(3).Caption
i = Label5.Caption
Dim l As Currency: Dim m As Single: Dim n As Single
l = 1000 * Int(h): m = 100000 * j: n = 1000000 * k
i = i + l + m + n
If i <= 5000000000# Then
Label5.Caption = i
Else
Label5.Caption = 5000000000#
End If
'火事の計算
Randomize
If Rnd * 12 < 1 Then
ka = Fix(Rnd * 6): kb = Fix(Rnd * 6)
kdd = ka * 6 + kb: kd = kdd: Call Kaji: ke = 1
ElseIf ke > 0 And ke < 11 Then
For r = 0 To ke
kd = kdd + ke + 5 * r: Call Kaji
kd = kdd - ke - 5 * r: Call Kaji
Next r
ke = ke + 1
End If
End Sub

Private Sub VScroll1_Change()
a = VScroll1.Value
Line2.Y1 = 1200 + 480 * VScroll1.Value
Line2.Y2 = 1200 + 480 * VScroll1.Value
End Sub
