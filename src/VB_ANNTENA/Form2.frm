VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "パラボラアンテナ"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows の既定値
   Begin VB.VScrollBar VScroll1 
      Height          =   3975
      Left            =   9960
      Max             =   100
      Min             =   10
      TabIndex        =   0
      Top             =   4080
      Value           =   10
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
s = 60
For a = 0 To 100 'OK
b = 100 - (a / VScroll1.Value) ^ 2
If a <> 0 Then 'OK
xx = -(a - oa) / (b - ob) 'OK
x = 1 / Tan(2 * Atn((a - oa) / (b - ob)))
aa = (a + oa) / 2
bb = (b + ob) / 2
If a < 0 Then Form1.Line ((aa - 100) * s, (bb - 100 * xx) * s)-((aa + 0) * s, (bb + 0 * xx) * s), RGB(0, 255, 0) 'OK
Form2.Line (aa * s, bb * s)-(aa * s, 0 * s), RGB(255, 255, 255)
Form2.Line ((aa - 100) * s, (bb - 100 * x) * s)-((aa + 0) * s, (bb + 0 * x) * s), RGB(255, 55 + 2 * a, 0) 'OK
Form2.Line (oa * s, ob * s)-(a * s, b * s), RGB(0, 0, 255) 'OK
End If 'OK
oa = a 'OK
ob = b 'OK
Next a
End Sub

Private Sub VScroll1_Change()
Form2.Cls
Call Form_Load
End Sub
