VERSION 5.00
Begin VB.UserControl ColorButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const default_Caption = "ColorButton"
Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Dim default_Font As StdFont

Dim m_Caption As String
Dim m_hasFocus As Boolean

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

' ColorButton Specific
Dim m_Mouse As Boolean

Public Property Let Caption(ByVal new_Caption As String)
    m_Caption = new_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    UserControl.BackColor = new_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    UserControl.ForeColor = new_ForeColor
    UserControl.Refresh
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Set Font(ByRef new_Font As StdFont)
    Set UserControl.Font = new_Font
    UserControl.Refresh
    PropertyChanged "Font"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Private Function getDefaultFont() As StdFont
    If default_Font Is Nothing Then
        Set getDefaultFont = Ambient.Font
    Else
        Set getDefaultFont = default_Font
    End If
End Function

Private Sub UserControl_Initialize()
    m_Mouse = False
    m_hasFocus = False
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = default_Caption
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    Set default_Font = UserControl.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", default_Caption)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", getDefaultFont())
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, default_Caption)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, default_BackColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, default_ForeColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, getDefaultFont())
End Sub

Sub notifyLeftButton(ByVal state As Boolean)
    If m_Mouse <> state Then
        m_Mouse = state
        Call UserControl.Refresh
        If Not m_Mouse Then RaiseEvent Click
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then notifyLeftButton True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then notifyLeftButton False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyCode)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_DblClick()
    notifyLeftButton True
End Sub

Sub updateFocus(ByVal state As Boolean)
    If m_hasFocus <> state Then
        m_hasFocus = state
        UserControl.Refresh
    End If
End Sub

Private Sub UserControl_GotFocus()
    updateFocus True
End Sub

Private Sub UserControl_LostFocus()
    updateFocus False
End Sub

Private Sub UserControl_Paint()
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If m_Mouse Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    UserControl.Print m_Caption

    If m_Mouse Then
        Call Graphics.DrawBorder(Me, ButtonPressed, 0, 0, w, h)
        Call Graphics.DrawBorder(Me, ButtonFocus, 0, 0, w, h)
    ElseIf m_hasFocus Then
        Call Graphics.DrawBorder(Me, ButtonOutsetBold, 0, 0, w, h)
        Call Graphics.DrawBorder(Me, ButtonFocus, 0, 0, w, h)
    Else
        Call Graphics.DrawBorder(Me, ButtonOutset, 0, 0, w, h)
    End If
End Sub

