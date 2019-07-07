VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'ŒÅ’èÂ°Ù ³¨ÝÄÞ³
   Caption         =   "“¹‹ï"
   ClientHeight    =   3930
   ClientLeft      =   6780
   ClientTop       =   2235
   ClientWidth     =   1695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mnset 
      Caption         =   "Ý’è"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mosdwn(3)

Private Sub Form_Load()
For a = 0 To 6
Line (a * 240, 240)-(a * 240, 13 * 240 - 15), RGB(239, 239, 239)
Next a
For b = 1 To 12
Line (0, b * 240)-(7 * 240 - 15, b * 240), RGB(255, 255, 255)
Next b
For c = 0 To 6
Line (c * 240 + 225, 240)-(c * 240 + 225, 13 * 240), RGB(95, 95, 95)
Next c
For d = 1 To 12
Line (0, d * 240 + 225)-(7 * 240, d * 240 + 225), RGB(63, 63, 63)
Next d
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And X < 240 * 7 And Y >= 240 And Y < 240 * 13 And Button = 1 Then
xx0 = Int(X / 240) * 240
yy0 = Int(Y / 240) * 240
xx1 = xx0 + 225
yy1 = yy0 + 225
Line (xx0, yy0)-(xx0, yy1), RGB(95, 95, 95)
Line (xx0, yy0)-(xx1, yy0), RGB(63, 63, 63)
Line (xx1, yy0)-(xx1, yy1), RGB(239, 239, 239)
Line (xx0, yy1)-(xx1, yy1), RGB(255, 255, 255)
mosdwn(0) = xx0
mosdwn(1) = yy0
mosdwn(2) = xx1
mosdwn(3) = yy1
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line (mosdwn(0), mosdwn(1))-(mosdwn(0), mosdwn(3)), RGB(239, 239, 239)
Line (mosdwn(0), mosdwn(1))-(mosdwn(2), mosdwn(1)), RGB(255, 255, 255)
Line (mosdwn(2), mosdwn(1))-(mosdwn(2), mosdwn(3)), RGB(95, 95, 95)
Line (mosdwn(0), mosdwn(3))-(mosdwn(2), mosdwn(3)), RGB(63, 63, 63)
If X >= 0 And X < 240 * 7 And Y >= 240 And Y < 240 * 13 And Button = 1 Then
xx0 = Int(X / 240) * 240
yy0 = Int(Y / 240) * 240
If mosdwn(0) = xx0 And mosdwn(1) = yy0 Then
Call buttonDWN(mosdwn(0) / 240, mosdwn(1) / 240 - 1)
End If
End If
End Sub

Public Sub buttonDWN(xx As Integer, yy As Integer)
Label1.Caption = xx & "   " & yy
End Sub

Private Sub mnset_Click()
Form3.Show
End Sub
