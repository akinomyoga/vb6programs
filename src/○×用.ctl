VERSION 5.00
Begin VB.UserControl ÅõÅ~ox 
   BackColor       =   &H0000FF00&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ScaleHeight     =   585.714
   ScaleMode       =   0  '’∞ªﬁ∞
   ScaleWidth      =   585.714
End
Attribute VB_Name = "ÅõÅ~ox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim a As Integer

Private Sub UserControl_Click()
a = a + 1: If a = 3 Then a = 0
If a = 0 Then BackColor = RGB(0, 255, 0)
If a = 1 Then BackColor = RGB(0, 0, 0)
If a = 2 Then BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserControl_Paint()
a = 0
End Sub
