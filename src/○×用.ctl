VERSION 5.00
Begin VB.UserControl ÅõÅ~ox 
   BackColor       =   &H0000FF00&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ScaleHeight     =   585.714
   ScaleMode       =   0  'User
   ScaleWidth      =   585.714
End
Attribute VB_Name = "ÅõÅ~ox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim a As Integer
Dim m_hover As Boolean

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Sub Rotate()
    a = a + 1: If a = 3 Then a = 0
    If a = 0 Then BackColor = RGB(0, 255, 0)
    If a = 1 Then BackColor = RGB(0, 0, 0)
    If a = 2 Then BackColor = RGB(255, 255, 255)
End Sub

Sub Reset()
    a = 0
End Sub

Sub updateHover(ByVal X As Single, ByVal Y As Single)
    m_hover = 0 <= X And X < ScaleWidth And 0 <= Y And Y < ScaleHeight
End Sub

Private Sub UserControl_DblClick()
    SetCapture hWnd
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    updateHover X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    updateHover X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    updateHover X, Y
    If Button = MouseButtonConstants.vbLeftButton And m_hover Then Rotate
End Sub

