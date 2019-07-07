VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "‹…‚Ì“à‘¤"
   ClientHeight    =   7845
   ClientLeft      =   1920
   ClientTop       =   1785
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   10095
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Show
s = 60
For a = 0 To 100 'OK
b = Sqr(100 ^ 2 - a ^ 2) 'OK
If a <> 0 Then 'OK
xx = -(a - oa) / (b - ob) 'OK
x = 1 / Tan(2 * Atn((a - oa) / (b - ob)))
aa = (a + oa) / 2
bb = (b + ob) / 2
If a < 0 Then Form1.Line ((aa - 100) * s, (bb - 100 * xx) * s)-((aa + 0) * s, (bb + 0 * xx) * s), RGB(0, 255, 0) 'OK
Form1.Line (aa * s, bb * s)-(aa * s, 0 * s), RGB(255, 255, 255)
Form1.Line ((aa - 100) * s, (bb - 100 * x) * s)-((aa + 0) * s, (bb + 0 * x) * s), RGB(255, 55 + 2 * a, 0) 'OK
Form1.Line (oa * s, ob * s)-(a * s, b * s), RGB(0, 0, 255) 'OK
End If 'OK
oa = a 'OK
ob = b 'OK
Next a
End Sub 'OK
