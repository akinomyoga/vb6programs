VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "分子模型"
   ClientHeight    =   6210
   ClientLeft      =   5640
   ClientTop       =   3030
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9165
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   2
   End
   Begin VB.Menu 入力 
      Caption         =   "入力..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim atodat(99, 6) As Double '原子の番号・三次元中心座標/色光三原色/半径,
Public atonum '原子の数
Dim two(2, 1), thrd(2, 2)
Dim win(500, 500, 2)
Dim xsee, ysee, zsee '視点
Dim pointcount As Single
Dim pc(0) As Single

Private Sub Form_Load()
Form3.Show
xsee = -1000
ysee = -1000
zsee = -1000
two(0, 0) = -2 * Sqr(5) / 5
two(0, 1) = -Sqr(5) / 5
two(1, 0) = 2 * Sqr(5) / 5
two(1, 1) = 2 * Sqr(5) / 5
two(2, 0) = 0
two(2, 1) = 1
atodat(0, 0) = 100
atodat(0, 4) = 255
atodat(0, 5) = 20
atodat(0, 6) = 20
For a = 0 To 500
For b = 0 To 500
win(a, b, 1) = 10000
Next b
Next a
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub 入力_Click()
Form2.Show
End Sub

Public Sub 表示()
Dim old(400, 2), old2(2)
If atonum > -1 Then
For c = 0 To atonum
r = atodat(c, 0)
stx = atodat(c, 1)
sty = atodat(c, 2)
stz = atodat(c, 3)
strd = atodat(c, 4)
stgr = atodat(c, 5)
stbl = atodat(c, 6)
For a = 0 To 2 Step 0.05
aa = a * 3.1415926535
zz = Sin(aa) * r
xy = Cos(aa) * r
For b = 0 To 1 Step 0.05
bb = b * 3.1415926535
xx = Cos(bb) * xy
yy = Sin(bb) * xy
If a > 0 And b > 0 Then
red = ((75 / 256 - a * 150 / 255) + strd / 256) * 256
If red > 255 Then red = 255
If red < 0 Then red = 0
gre = ((75 / 256 - a * 150 / 255) + stgr / 256) * 256
If gre > 255 Then gre = 255
If gre < 0 Then gre = 0
blu = ((75 / 256 - a * 150 / 255) + stbl / 256) * 256
If blu > 255 Then blu = 255
If blu < 0 Then blu = 0
Call draw3dline(xx + stx, yy + sty, zz + stz, xo + stx, yo + sty, zo + stz, RGB(red, gre, blu))
Call draw3dline(xx + stx, yy + sty, zz + stz, old(b * 200 - 1, 0) + stx, old(b * 200 - 1, 0) + sty, old(b * 200 - 1, 0) + stz, RGB(red, gre, blu))
End If
old(b * 200, 0) = xx
old(b * 200, 1) = yy
old(b * 200, 2) = zz
xo = xx
yo = xx
zo = xx
Next b
ProgressBar1.Value = a
Next a
Next c
End If
Call showing
End Sub

Public Sub draw3dline(bx, by, bz, ex, ey, ez, co)
bx2 = thrd(0, 0) * bx + thrd(1, 0) * by + thrd(2, 0) * bz
by2 = thrd(0, 1) * bx + thrd(1, 1) * by + thrd(2, 1) * bz
bz2 = thrd(0, 2) * bx + thrd(1, 2) * by + thrd(2, 2) * bz
ex2 = thrd(0, 0) * ex + thrd(1, 0) * ey + thrd(2, 0) * ez
ey2 = thrd(0, 1) * ex + thrd(1, 1) * ey + thrd(2, 1) * ez
ez2 = thrd(0, 2) * ex + thrd(1, 2) * ey + thrd(2, 2) * ez
xdis = xsee - (bx2 + ex2) / 2
ydis = ysee - (by2 + ey2) / 2
zdis = zsee - (bz2 + ez2) / 2
dist = Sqr(xdis ^ 2 + ydis ^ 2 + zdis ^ 2)
sx = two(0, 0) * bx + two(1, 0) * by + two(2, 0) * by
sy = two(0, 1) * bx + two(1, 1) * by + two(2, 1) * by
fx = two(0, 0) * ex + two(1, 0) * ey + two(2, 0) * ey
fy = two(0, 1) * ex + two(1, 1) * ey + two(2, 1) * ey
leng = Abs(sx - sy)
If leng < Abs(sy - fy) Then leng = Abs(sy - fy)
If leng = 0 Then leng = 1
For n = 0 To leng
dx = (sx - fx) / leng * n + fx + 250
dy = (sy - fy) / leng * n + fy + 250
If dx >= 0 And dx <= 500 And dy >= 0 And dy <= 500 Then
If win(dx, dy, 1) >= dist Then
win(500, 500, 1) = dist
If win(dx, dy, 0) <> co Then
win(dx, dy, 0) = co
win(dx, dy, 2) = 1
pointcount = pointcount + 1
pc(0) = (pc(0) + co) / 2
End If
End If
End If
Next n
End Sub

Public Sub showing()
For a = 0 To 500
For b = 0 To 500
If win(a, b, 2) = 1 Then
win(a, b, 2) = 0
Form1.PSet (a * 15 + 300, b * 15 + 300), win(a, b, 0)
End If
Next b
Next a
MsgBox "表示"
MsgBox pointcount & "," & pc(0)
End Sub

Public Sub 移動thrd(a, b, c, d, e, f, g, h, i)
thrd(0, 0) = a
thrd(0, 1) = b
thrd(0, 2) = c
thrd(1, 0) = d
thrd(1, 1) = e
thrd(1, 2) = f
thrd(2, 0) = g
thrd(2, 1) = h
thrd(2, 2) = i
End Sub
