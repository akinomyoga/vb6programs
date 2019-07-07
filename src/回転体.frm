VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "âÒì]ëÃ"
   ClientHeight    =   9900
   ClientLeft      =   765
   ClientTop       =   900
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "âÒì]ëÃ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13965
   Begin VB.VScrollBar VScroll3 
      Height          =   2175
      Left            =   11520
      Max             =   360
      SmallChange     =   5
      TabIndex        =   16
      Top             =   6960
      Width           =   495
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   2175
      Left            =   12240
      Max             =   360
      SmallChange     =   5
      TabIndex        =   15
      Top             =   6960
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "éÂÇ»ê¸Çï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   5520
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÇÌÇ¡Ç©Ç‡ï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2175
      Left            =   10920
      Max             =   360
      SmallChange     =   5
      TabIndex        =   11
      Top             =   6960
      Value           =   30
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   10320
      Max             =   12
      Min             =   1
      TabIndex        =   9
      Top             =   6480
      Value           =   10
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "è¡Ç∑"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "èIóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   5
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "âÒì]ëÃÇÃï\é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ç‚ÇËíºÇ∑"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   10200
      ScaleHeight     =   3435
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "å©ÇÈäpìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "10ìxíuÇ´Ç…ï`Ç´Ç‹Ç∑"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   10
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   11880
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ñ{"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   3
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sx As Integer, sy As Integer, bx(49) As Integer, by(49) As Integer, fx(49) As Integer, fy(49) As Integer
Dim lc As Integer, lo As Integer, mo As Integer, no As Integer, po As Integer
Dim st(2, 360, 49) As Integer, en(2, 360, 49) As Integer
Dim sina As Double, sinb As Double, sinc As Double, cosa As Double, cosb As Double, cosc As Double
Private Sub Command1_Click()
Picture1.Cls
lc = 0
Label1.Caption = 0
End Sub

Private Sub Command2_Click()
Dim kosu As Integer
kosu = 360 / HScroll1.Value
If lc <> 0 Then
For ld = 0 To lc - 1

For oo = 0 To kosu
o = oo * HScroll1.Value
Dim l As Integer, m As Integer, n As Integer, p As Integer
st(0, oo, ld) = Int(bx(ld) * Cos(o * 3.1415 / 180))
st(1, oo, ld) = Int(by(ld))
st(2, oo, ld) = bx(ld) * Sin(o * 3.1415 / 180)
en(0, oo, ld) = Int(fx(ld) * Cos(o * 3.1415 / 180))
en(1, oo, ld) = Int(fy(ld))
en(2, oo, ld) = fx(ld) * Sin(o * 3.1415 / 180)
Next oo

For oo = 0 To kosu
qx = st(0, oo, ld)
qy = st(1, oo, ld)
qz = st(2, oo, ld)
qqy = qy * cosa - qz * sina
qqz = qy * sina + qz * cosa
qy = qqy
qz = qqz
qqx = qx * cosc - qz * sinc
qqz = qx * sinc + qz * cosc
qx = qqx
qz = qqz
qqx = qx * cosb - qy * sinb
qqy = qx * sinb + qy * cosb
qx = qqx
qy = qqy

rx = en(0, oo, ld)
ry = en(1, oo, ld)
rz = en(2, oo, ld)
rry = ry * cosa - rz * sina
rrz = ry * sina + rz * cosa
ry = rry
rz = rrz
rrx = rx * cosc - rz * sinc
rrz = rx * sinc + rz * cosc
rx = rrx
rz = rrz
rrx = rx * cosb - ry * sinb
rry = rx * sinb + ry * cosb
rx = rrx
ry = rry

l = qx
m = qy
n = rx
p = ry
If Check2.Value = 1 Then Line (l + 4980, m + 4980)-(n + 4980, p + 4980)
If oo <> 0 And Check1.Value = 1 Then
Line (lo + 4980, mo + 4980)-(l + 4980, m + 4980)
Line (no + 4980, po + 4980)-(n + 4980, p + 4980)
End If
lo = l
mo = m
no = n
po = p
Next oo
Next ld
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Cls
End Sub

Private Sub Form_Load()
Call VScroll1_Change
Call VScroll2_Change
Call VScroll3_Change
End Sub

Private Sub HScroll1_Change()
Label5.Caption = HScroll1.Value & "ìxíuÇ´Ç…ï`Ç´Ç‹Ç∑"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
sx = X
sy = Y
Label4.Caption = sx & "-" & sy
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = X & "-" & Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lc >= 0 And lc <= 49 Then
Picture1.DrawWidth = 1
Picture1.DrawStyle = 0
Picture1.Line (sx, sy)-(X, Y), QBColor(10)
bx(lc) = sx
by(lc) = sy
fx(lc) = X
fy(lc) = Y
lc = lc + 1
Label1.Caption = lc
End If
Label4.Caption = ""
End Sub

Private Sub VScroll1_Change()
sina = Sin(VScroll1.Value * 3.1415 / 180)
cosa = Cos(VScroll1.Value * 3.1415 / 180)
Cls
Call Command2_Click
End Sub

Private Sub VScroll2_Change()
sinb = Sin(VScroll2.Value * 3.1415 / 180)
cosb = Cos(VScroll2.Value * 3.1415 / 180)
Cls
Call Command2_Click
End Sub

Private Sub VScroll3_Change()
sinc = Sin(VScroll3.Value * 3.1415 / 180)
cosc = Cos(VScroll3.Value * 3.1415 / 180)
Cls
Call Command2_Click
End Sub
