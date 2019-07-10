VERSION 5.00
Begin VB.UserControl ToggleButton 
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ToolboxBitmap   =   "ToggleButton.ctx":0000
End
Attribute VB_Name = "ToggleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_hasFocus As Boolean
Dim m_leftButton As Boolean
Dim m_hover As Boolean

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_Caption = "ToggleButton"
Const default_Value = False
Dim m_Caption As String
Dim m_Value As Boolean

Public Event Click()

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Dim default_Font As StdFont

Public Event MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Boolean)
    m_Value = new_Value
    UserControl.Refresh
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_Caption As String)
    m_Caption = new_Caption
    PropertyChanged "Caption"
End Property

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    UserControl.BackColor = new_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    UserControl.ForeColor = new_ForeColor
    UserControl.Refresh
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef new_Font As StdFont)
    Set UserControl.Font = new_Font
    UserControl.Refresh
    PropertyChanged "Font"
End Property

Private Function getDefaultFont() As StdFont
    If default_Font Is Nothing Then
        Set getDefaultFont = Ambient.Font
    Else
        Set getDefaultFont = default_Font
    End If
End Function

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Sub toggleState()
    If m_Value Then
        m_Value = False
    Else
        m_Value = True
    End If
    RaiseEvent Click
    UserControl.Refresh
End Sub

Sub notifyLeftButton(ByVal state As Boolean)
    If m_leftButton <> state Then
        m_leftButton = state
        If m_hover And Not state Then
            Call toggleState
        Else
            Call UserControl.Refresh
        End If
    End If
End Sub

Sub updateFocus(ByVal state As Boolean)
    If m_hasFocus <> state Then
        m_hasFocus = state
        UserControl.Refresh
    End If
End Sub

Sub hover_Update(ByVal X As Single, ByVal Y As Single)
    oldHover = m_hover
    m_hover = 0 <= X And X < ScaleWidth And 0 <= Y And Y < ScaleHeight
    If m_leftButton And m_hover <> oldHover Then
        UserControl.Refresh
    End If
End Sub

Sub onPaint()
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If m_leftButton And m_hover Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    UserControl.Print m_Caption

    If m_leftButton And m_hover Then
        If m_hasFocus Then
            Call Graphics.DrawBorder(Me, kbBorderButtonInset, 0, 0, w, h)
            Call Graphics.DrawBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call Graphics.DrawBorder(Me, kbBorderButtonInset, 0, 0, w, h)
            If Value Then UserControl.Line (4, 4)-(w - 5, h - 5), SystemColorConstants.vb3DDKShadow, B
        End If
    ElseIf Value Then
        If m_hasFocus Then
            Call Graphics.DrawBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            Call Graphics.DrawBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call Graphics.DrawBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            UserControl.Line (4, 4)-(w - 5, h - 5), SystemColorConstants.vb3DDKShadow, B
        End If
    Else
        If m_hasFocus Then
            Call Graphics.DrawBorder(Me, kbBorderButtonOutsetBold, 0, 0, w, h)
            Call Graphics.DrawBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call Graphics.DrawBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        End If
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    m_leftButton = False
    m_hasFocus = False
    m_Value = default_Value
    m_Caption = default_Caption
End Sub

Private Sub UserControl_InitProperties()
    m_Value = default_Value
    m_Caption = default_Caption
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    If default_Font Is Nothing Then
        Set default_Font = UserControl.Font
    End If
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

' イベントは MouseDown, MouseUp, Click / DblClick, MouseUp の順で発生するそうだ。
' http://cya.sakura.ne.jp/vb/MSHFlexGrid_Event.htm
Private Sub UserControl_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    hover_Update X, Y
    If button = vbLeftButton Then notifyLeftButton True
    RaiseEvent MouseDown(button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    Graphics.ReleaseCapture
    If button = vbLeftButton Then notifyLeftButton False
    hover_Update X, Y
    RaiseEvent MouseUp(button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    hover_Update X, Y
    RaiseEvent MouseMove(button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_DblClick()
    notifyLeftButton True
    Graphics.SetCapture UserControl.hWnd
End Sub

Private Sub UserControl_GotFocus()
    updateFocus True
End Sub

Private Sub UserControl_LostFocus()
    updateFocus False
End Sub

Private Sub UserControl_Paint()
    onPaint
End Sub

