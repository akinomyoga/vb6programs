VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'ŒÅ’è(ŽÀü)
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   5115
   ClientTop       =   3615
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5340
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "life4.frx":0000
      Left            =   3600
      List            =   "life4.frx":0010
      TabIndex        =   2
      Text            =   "pen"
      Top             =   120
      Width           =   1575
   End
   Begin life4.SpinText spintext1 
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin life4.SpinText spintext1 
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin life4.SpinText spintext1 
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      _ExtentX        =   3201
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dnmouse(3), lifedat(200, 200) As Byte

Private Sub Combo1_Change()
Select Case Combo1.Text
Case "pen"
Form1.MousePointer = 0
Case "line", "square", "circle"
Form1.MousePointer = 2
End Select
End Sub

Private Sub Form_Load()
spintext1(0).LabelCaption = "x"
spintext1(1).LabelCaption = "y"
spintext1(2).LabelCaption = "z"
For a = 0 To 2
With spintext1(a)
.SpinMax = 200
.SpinMin = 10
.SpinValue = 100
End With
Next a
For a = 0 To 200
Line (0, a * 15)-(3015, a * 15), &HFFFFFF
Next a
Line (0, 3015)-(3015, 3015)
Line (3015, 0)-(3015, 3030)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 Line (x - 120, y)-(x + 120, y), &HFF00
 Line (x, y - 120)-(x, y + 120), &HFF00
End If
dnmouse(0) = Button
dnmouse(1) = Shift
dnmouse(2) = Int(x / 15)
dnmouse(3) = Int(y / 15)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 xx = Int(x / 15)
 yy = Int(y / 15)
 Select Case Combo1.Text
  Case "pen"
   Call set_lifedat(xx, yy)
  Case "line"
   aa0 = dnmouse(2) - xx
   aa1 = dnmouse(3) - yy
   If Abs(aa0) > Abs(aa1) Then
    stp = 1: If aa0 < 0 Then stp = -1
    For a = 0 To aa0 Step stp
     Call set_lifedat(a + xx, Int(aa1 / aa0 * a + 0.5) + yy)
    Next a
   Else
    stp = 1: If aa1 < 0 Then stp = -1
    For a = 0 To aa1 Step stp
     Call set_lifedat(Int(aa0 / aa1 * a + 0.5) + xx, a + yy)
    Next a
   End If
  Case "square"
   If dnmouse(2) > xx Then
    x0 = xx
    X1 = dnmouse(2)
   Else
    x0 = dnmouse(2)
    X1 = xx
   End If
   If dnmouse(3) > yy Then
    y0 = yy
    Y1 = dnmouse(3)
   Else
    y0 = dnmouse(3)
    Y1 = yy
   End If
   For xa = x0 To X1
    Call set_lifedat(xa, y0)
    Call set_lifedat(xa, Y1)
   Next xa
   For ya = y0 To Y1
    Call set_lifedat(x0, ya)
    Call set_lifedat(X1, ya)
   Next ya
  Case "circle"
   xo = Int((dnmouse(2) + xx) / 2)
   yo = Int((dnmouse(3) + xx) / 2)
   xr = Abs(dnmouse(2) - xo)
   yr = Abs(dnmouse(3) - yo)
   xrr = xr * xr
   yrr = yr * yr
   yxr = yr / xr
   xyr = xr / yr
   For xa = 0 To xr
    y0 = Int(Sqr(xrr - xa * xa) * yxr)
    Call set_lifedat(xo + xa, yo + y0)
    Call set_lifedat(xo - xa, yo + y0)
    Call set_lifedat(xo + xa, yo - y0)
    Call set_lifedat(xo - xa, yo - y0)
   Next xa
   For ya = 0 To yr
    x0 = Int(Sqr(yrr - ya * ya) * xyr)
    Call set_lifedat(xo + x0, yo + ya)
    Call set_lifedat(xo - x0, yo + ya)
    Call set_lifedat(xo + x0, yo - ya)
    Call set_lifedat(xo - x0, yo - ya)
   Next ya
 End Select
 Call drawB(dnmouse(2), dnmouse(3))
End If
End Sub

Private Sub set_lifedat(x, y)
On Error GoTo err
Dim xx As Integer, yy As Integer
xx = Int(x)
yy = Int(y)
If 0 <= xx And xx <= spintext1(0).SpinValue And 0 <= yy And yy <= spintext1(1).SpinValue Then
 lifedat(xx, yy) = 1
 Form1.PSet (xx * 15, yy * 15), 0
End If
err:
End Sub

Private Sub drawA()
 For a = 0 To spintext1(0).SpinValue
  For b = 0 To spintext1(1).SpinValue
   If lifedat(a, b) = 1 Then
    Form1.PSet (a * 15, b * 15)
   Else
    Form1.PSet (a * 15, b * 15), &HFFFFFF
   End If
  Next b
 Next a
End Sub

Private Sub drawB(x, y)
 For a = -8 To 8
  If 0 <= x + a And x + a <= 200 And 0 <= y And y <= 200 Then
   If lifedat(x + a, y) = 1 Then co = 0 Else co = &HFFFFFF
   Form1.PSet ((x + a) * 15, y * 15), co
  End If
  If 0 <= y + a And y + a <= 200 And 0 <= x And x <= 200 Then
  If lifedat(x, y + a) = 1 Then co = 0 Else co = &HFFFFFF
  Form1.PSet (x * 15, (y + a) * 15), co
  End If
 Next a
End Sub
