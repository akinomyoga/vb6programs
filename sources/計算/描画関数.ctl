VERSION 5.00
Begin VB.UserControl GraphF 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   Picture         =   "ï`âÊä÷êî.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "ï`âÊä÷êî.ctx":0342
End
Attribute VB_Name = "GraphF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type BitDat
C As Long
d As Double
End Type

Private Type xyc
x As Integer
Y As Integer
C As Long
End Type

Private Type Shapes
n As Byte
dat(99) As xyc
End Type

Private Type Patterns
xn As Byte
Yn As Byte
x_rvs As Byte
y_rvs As Byte
dat(99, 99) As Long
mask(99, 99) As Byte
End Type

Private ptn As Patterns, ptn1(6) As Patterns
Private pict(999, 999) As BitDat, shp(99) As Shapes, clr(99) As Long
Private ptns(99) As Patterns
Public BorderStyleP As Integer, FillStyleP As Integer, FillWidthP As Integer, PatternNo As Byte
'borderstylep 0=é¿ê¸ 1=çΩê¸ 2=àÍì_çΩê¸ 3=ìÒì_çΩê¸ 4=ì_ê¸
'fillstylep 0=ìßñæ 1=ìhÇËÇ¬Ç‘Çµ 2=èc 3=â° 4=âEâ∫ 5=ç∂â∫ 6=ècâ° 7=éŒñ‘ 8=ñÕól(ptns(PatternNo) éQè∆)
'shape **[ì_ÇÃêî](***[x]***[y]***[r]***[g]***[b])Å~n(ì_ÇÃêî)

Private Sub UserControl_Initialize()
FillWidthP = 7
For a = 0 To 99
For B = 0 To 99
For C = 0 To 6
ptn1(C).mask(a, B) = 255
Next C
Next B
Next a
For a = 0 To 99
ptn1(2).mask(0, a) = 2
ptn1(3).mask(a, 0) = 2
ptn1(4).mask(a, a) = 2
ptn1(5).mask(a, a) = 2
ptn1(6).mask(0, a) = 2
ptn1(6).mask(a, 0) = 2
Next a
ptn1(5).x_rvs = 1
End Sub

Public Sub Set_shape(number As Integer, shape)
On Error GoTo ErrH
If 0 <= number And number <= 99 Then
For n = 0 To number
shp(n).n = Left(shape, 2)
For a = 0 To Left(shape, 2)
shp(n).dat(a).x = Mid(shape, 15 * a + 3, 3)
shp(n).dat(a).Y = Mid(shape, 15 * a + 6, 3)
shp(n).dat(a).C = RGB(Mid(shape, 15 * a + 9, 3), Mid(shape, 15 * a + 12, 3), Mid(shape, 15 * a + 15, 3))
Next a
Next n
End If
Exit Sub
ErrH:
MsgBox "ñ‚ëËÇ™ãNÇ±ÇËÇ‹ÇµÇΩÅBSet_shapeÃﬂ€º∞ºﬁ¨ÇÕíÜífÇµÇ‹Ç∑ÅB"
End Sub

Public Sub PointP(x As Integer, Y As Integer, color0 As Long, shpNo)
Select Case shpNo
Case 0 To 99
For a = 0 To shp(shpNo).n
x0 = x + shp(shpNo).dat(a).x
y0 = Y + shp(shpNo).dat(a).Y
If (0 <= x0 And x0 <= 999) And (0 <= y0 And y0 <= 999) Then pict(x0, y0).C = shp(shpNo).dat(a).C
Next a
Case Else
If (0 <= x And x <= 999) And (0 <= Y And Y <= 999) Then pict(x, Y).C = color0
End Select
End Sub

Public Sub LineP(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, shpNo)
Dim color0 As Long, xp As Integer, yp As Integer
color0 = clr(0)
xx = x1 - x2
yy = y1 - y2
If Abs(xx) > Abs(yy) Then
For x = 0 To Abs(xx)
modd = Int((x / 4) Mod 8)
Select Case BorderStyleP
Case 1: If modd = 6 Or modd = 7 Then GoTo Skip1
Case 2
If modd = 4 Or modd = 6 Or modd = 7 Then GoTo Skip1
Case 3
If modd = 2 Or modd = 4 Or modd = 6 Or modd = 7 Then GoTo Skip1
Case 4: If modd Mod 2 = 1 Then GoTo Skip1
End Select
xp = x1 + x
yp = y1 + Int(yy / xx * x)
Call PointP(xp, yp, color0, shpNo)
Skip1:
Next x
Else
For Y = y1 To y2
xp = x1 + Int(xx / yy * Y)
yp = y1 + Y
Call PointP(xp, yp, color0, shpNo)
Next Y
End If
End Sub

Public Sub SquareP(x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer, color0 As Long, shpNo, SclMvmnt)
Dim x0 As Integer, y0 As Integer
If x1 > x2 Then C = x1: x1 = x2: x2 = C
If y1 > y2 Then C = y1: y1 = y2: y2 = C
If SclMvmnt Then x0 = x1: y0 = y1 Else x0 = 0: y0 = 0
If Abs(x1 - x2) > 1 And Abs(y1 - y2) > 1 Then
Select Case FillStyleP
Case 1 To 5
ptn = ptn1(FillStyleP)
ptn.xn = FillWidthP
ptn.Yn = FillWidthP
GoSub Filling
Case 6
ptn = ptn1(3)
ptn.xn = FillWidthP
ptn.Yn = FillWidthP
GoSub Filling
ptn = ptn1(4)
ptn.xn = FillWidthP
ptn.Yn = FillWidthP
GoSub Filling
Case 8
If PatternNo <= 99 Then
ptn = ptns(PatternNo)
ptn.xn = FillWidthP
ptn.Yn = FillWidthP
GoSub Filling
End If
End Select
End If
Call LineP(x1, y1, x1, y2, shpNo)
Call LineP(x2, y1, x2, y2, shpNo)
Call LineP(x1, y1, x2, y1, shpNo)
Call LineP(x1, y2, x2, y2, shpNo)
Exit Sub
Filling:
For a = x1 + 1 To x2 - 1
For B = y1 + 1 To y2 - 1
Call PatternP(x0, y0, (a), (B))
Next B
Next a
Return
End Sub

Private Sub PatternP(x0 As Integer, y0 As Integer, x As Integer, Y As Integer)
x1 = (x - x0) Mod (ptn.xn + 1): If x1 < 0 Then x1 = x1 + ptn.xn + 1: If ptn.x_rvs Then x1 = ptn.xn - x1
y1 = (Y - y0) Mod (ptn.Yn + 1): If y1 < 0 Then y1 = y1 + ptn.Yn + 1: If ptn.y_rvs Then yi = ptn.Yn - y1
If (0 <= x And x <= 999) And (0 <= Y And Y <= 999) Then
Select Case ptn.mask(x1, y1)
Case 0: pict(x, Y).C = ptn.dat(x1, y1)
Case 1 To 100: pict(x, Y).C = clr(ptn.mask(x1, y1) - 1)
End Select
End If
End Sub

Public Sub ClsP()
For x = 0 To 999
For Y = 0 To 999
pict(x, Y).d = 0
pict(x, Y).C = 0
Next Y
Next x
End Sub

Public Function PictF(x As Integer, Y As Integer, col_or_dis)
If x < 0 Then x = 0 Else: If x > 999 Then x = 999
If Y < 0 Then Y = 0 Else: If Y > 999 Then Y = 999
Select Case col_or_dis
Case 1, "d", "dis", "distance": pivtf = pict(x, Y).d
Case Else: PictF = pict(x, Y).C
End Select
End Function

Public Sub TriangleP(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, x3 As Integer, y3 As Integer, shpNo, SclMvmnt)
Call LineP(x1, y1, x2, y2, shpNo)
Call LineP(x1, y1, x3, y3, shpNo)
Call LineP(x2, y2, x3, y3, shpNo)
End Sub

