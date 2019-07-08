VERSION 5.00
Begin VB.UserControl ToggleButton 
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
End
Attribute VB_Name = "ToggleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const default_Value = False
Const default_Caption = "ToggleButton"
Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Dim default_Font As StdFont

Dim m_Value As Boolean
Dim m_Caption As String
Dim m_Mouse As Boolean

Public Event Click()

Public Property Let Value(ByVal new_Value As Boolean)
    m_Value = new_Value
    UserControl.Refresh
    PropertyChanged "Value"
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

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
End Sub

Private Sub UserControl_InitProperties()
    m_Value = default_Value
    m_Caption = default_Caption
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    Set default_Font = UserControl.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Caption = PropBag.ReadProperty("Caption", default_Caption)
  m_Value = PropBag.ReadProperty("Value", default_Value)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
  Set UserControl.Font = PropBag.ReadProperty("Font", getDefaultFont())
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Caption", m_Caption, default_Caption)
  Call PropBag.WriteProperty("Value", m_Value, default_Value)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, default_BackColor)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, default_ForeColor)
  Call PropBag.WriteProperty("Font", UserControl.Font, getDefaultFont())
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Mouse = True
    Call UserControl.Refresh
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Mouse = False
    Call UserControl.Refresh
End Sub

Private Sub ToggleState()
    If m_Value Then
        m_Value = False
    Else
        m_Value = True
    End If
    RaiseEvent Click
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    Call ToggleState
End Sub

Private Sub UserControl_DblClick()
    Call ToggleState
End Sub

Private Sub UserControl_Paint()
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    If m_Mouse Then
        UserControl.Line (0, 0)-(w - 1, h - 1), SystemColorConstants.vb3DShadow, BF
    End If
    
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If Value Or m_Mouse Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    UserControl.Print m_Caption
    
    If Value Or m_Mouse Then
        UserControl.Line (0, 0)-(w - 1, 0), SystemColorConstants.vb3DShadow
        UserControl.Line (0, 0)-(0, h - 1), SystemColorConstants.vb3DShadow
        UserControl.Line (1, 1)-(w - 1, 1), SystemColorConstants.vb3DDKShadow
        UserControl.Line (1, 1)-(1, h - 1), SystemColorConstants.vb3DDKShadow
        UserControl.Line (1, h - 2)-(w - 1, h - 2), SystemColorConstants.vb3DHighlight
        UserControl.Line (w - 2, 1)-(w - 2, h - 1), SystemColorConstants.vb3DHighlight
        UserControl.Line (0, h - 1)-(w, h - 1), SystemColorConstants.vb3DLight
        UserControl.Line (w - 1, 0)-(w - 1, h), SystemColorConstants.vb3DLight
    Else
        UserControl.Line (0, 0)-(w - 1, 0), SystemColorConstants.vb3DHighlight
        UserControl.Line (0, 0)-(0, h - 1), SystemColorConstants.vb3DHighlight
        UserControl.Line (1, 1)-(w - 1, 1), SystemColorConstants.vb3DLight
        UserControl.Line (1, 1)-(1, h - 1), SystemColorConstants.vb3DLight
        UserControl.Line (1, h - 2)-(w - 1, h - 2), SystemColorConstants.vb3DShadow
        UserControl.Line (w - 2, 1)-(w - 2, h - 1), SystemColorConstants.vb3DShadow
        UserControl.Line (0, h - 1)-(w, h - 1), SystemColorConstants.vb3DDKShadow
        UserControl.Line (w - 1, 0)-(w - 1, h), SystemColorConstants.vb3DDKShadow
    End If
End Sub

