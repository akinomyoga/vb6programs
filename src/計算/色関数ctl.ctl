VERSION 5.00
Begin VB.UserControl ColorF 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   Picture         =   "êFä÷êîctl.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "êFä÷êîctl.ctx":0342
End
Attribute VB_Name = "ColorF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function CMY(c As Byte, m As Byte, Y As Byte) As Long
CMY = RGB(255 - c, 255 - m, 255 - Y)
End Function

Public Function HSV(h As Integer, s As Byte, v As Byte) As Long
h = h Mod 360
If h < 0 Then h = h + 360
F = h Mod 60
p1 = v * (255 - s)
p2 = v * (255 * 255 - s * F)
p3 = v * (255 * 255 - s * (255 - F))
Select Case Int(h / 60)
Case 0
HSV = RGB(v, p3, p1)
Case 1
HSV = RGB(p2, v, p1)
Case 2
HSV = RGB(p1, v, p3)
Case 3
HSV = RGB(p1, p2, v)
Case 4
HSV = RGB(p3, p1, v)
Case 5
HSV = RGB(v, p1, p2)
End Select
End Function

Public Function HLS(h As Integer, l As Byte, s As Byte) As Long
h = h Mod 360
If h < 0 Then h = h + 360
End Function

Public Function HSL(h As Integer, s As Byte, l As Byte) As Long
h = h Mod 360
If h < 0 Then h = h + 360
End Function

Public Function CIE(x As Integer, Y As Integer, b As Integer) As Long

End Function

Public Function YIQ(Y As Integer, i As Integer, q As Integer) As Long
rr = Y + 0.956 * i + 0.62 * q
gg = Y - 0.272 * i - 0.647 * q
bb = Y - 1.108 * i + 1.705 * q
YIQ = RGB(rr, gg, bb)
End Function

Public Function YUB(Y As Integer, u As Integer, b As Integer) As Long
rr = Y + 1.14 * b
gg = Y - 0.395 * u - 0.581 * b
bb = Y + 2.032 * u
YUB = RGB(rr, gg, bb)
End Function

Public Function YCbCr(Y As Integer, Cb As Integer, Cr As Integer) As Long
rr = Y - 0.001 * Cb + 1.402 * Cr
gg = Y - 0.344 * Cb - 0.714 * Cr
bb = Y + 1.772 * Cb + 0.001 * Cr
YCbCr = RGB(rr, gg, bb)
End Function

Public Function YPbPr(Y As Integer, Cb As Integer, Cr As Integer) As Long
rr = -1.737 * Y + 0.001 * Pb + 2.737 * Pr
gg = 1.828 * Y - 0.277 * Pb - 0.828 * Pr
bb = 1.001 * Y + 1.826 * Pb - 0.001 * Pr
YPbPr = RGB(rr, gg, bb)
End Function

Public Function CMYeB(c As Byte, m As Byte, Ye As Byte, b As Byte) As Long
CMYeB = RGB(255 - c - b, 255 - m - b, 255 - Ye - b)
End Function
