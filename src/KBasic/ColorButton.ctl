VERSION 5.00
Begin VB.UserControl ColorButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ColorButton.ctx":0000
End
Attribute VB_Name = "ColorButton"
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

Const default_Caption = "ColorButton"

Dim m_Caption As String

Public Event Click()

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_Enabled = True
Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Dim default_Font As StdFont
Const default_Tag = ""
Const default_MousePointer = MousePointerConstants.vbDefault
Dim default_MouseIcon As IPictureDisp

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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_Caption As String)
    If m_Caption <> new_Caption Then
        m_Caption = new_Caption
        PropertyChanged "Caption"
    End If
End Property

Sub ownProperties_Initialize()
    m_Caption = default_Caption
End Sub

Sub ownProperties_Read(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", default_Caption)
End Sub

Sub ownProperties_Write(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, default_Caption)
End Sub

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
    If UserControl.Enabled <> new_Enabled Then
        UserControl.Enabled = new_Enabled
        If Not new_Enabled Then
            m_hasFocus = False
            m_leftButton = False
        End If
        UserControl.Refresh
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    If UserControl.BackColor <> BackColor Then
        UserControl.BackColor = new_BackColor
        UserControl.Refresh
        PropertyChanged "BackColor"
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    If UserControl.ForeColor <> new_ForeColor Then
        UserControl.ForeColor = new_ForeColor
        UserControl.Refresh
        PropertyChanged "ForeColor"
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef new_Font As StdFont)
    If UserControl.Font <> new_Font Then
        Set UserControl.Font = new_Font
        UserControl.Refresh
        PropertyChanged "Font"
    End If
End Property

Private Function getDefaultFont() As StdFont
    If default_Font Is Nothing Then
        Set getDefaultFont = Ambient.Font
    Else
        Set getDefaultFont = default_Font
    End If
End Function

Public Property Get Tag() As String
    Tag = UserControl.Tag
End Property

Public Property Let Tag(ByVal new_Tag As String)
    If UserControl.Tag <> new_Tag Then
        UserControl.Tag = new_Tag
        PropertyChanged "Tag"
    End If
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal new_MousePointer As Integer)
    If UserControl.MousePointer <> new_MousePointer Then
        UserControl.MousePointer = new_MousePointer
        PropertyChanged "MousePointer"
    End If
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByRef new_MouseIcon As IPictureDisp)
    If UserControl.MouseIcon <> new_MouseIcon Then
        Set UserControl.MouseIcon = new_MouseIcon
        PropertyChanged "MouseIcon"
    End If
End Property

Sub delegateProperties_ctor()
    Set default_MouseIcon = Nothing
End Sub

Sub delegateProperties_Initialize()
    UserControl.Enabled = default_Enabled
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    If default_Font Is Nothing Then Set default_Font = UserControl.Font
    UserControl.Tag = default_Tag
    UserControl.MousePointer = default_MousePointer
    Set UserControl.MouseIcon = default_MouseIcon
End Sub

Sub delegateProperties_Read(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", default_Enabled)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", getDefaultFont())
    UserControl.Tag = PropBag.ReadProperty("Tag", default_Tag)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", default_MousePointer)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", default_MouseIcon)
End Sub

Sub delegateProperties_Write(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, default_Enabled)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, default_BackColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, default_ForeColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, getDefaultFont())
    Call PropBag.WriteProperty("Tag", UserControl.Tag, default_Tag)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, default_MousePointer)
    Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, default_MouseIcon)
End Sub

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Sub notifyLeftButton(ByVal state As Boolean)
    If m_leftButton <> state Then
        m_leftButton = state
        Call UserControl.Refresh
        If Not m_leftButton And m_hover Then RaiseEvent Click
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
    
    If UserControl.Enabled Then
        If m_leftButton And m_hover Then
            CurrentX = CurrentX + 1
            CurrentY = CurrentY + 1
        End If
        UserControl.Print m_Caption
        
        If m_leftButton And m_hover Then
            Call Graphics.DrawBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            Call Graphics.DrawBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        ElseIf m_hasFocus Then
            Call Graphics.DrawBorder(Me, kbBorderButtonOutsetBold, 0, 0, w, h)
            Call Graphics.DrawBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call Graphics.DrawBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        End If
    Else
        oldForeColor = UserControl.ForeColor
        x0 = CurrentX
        y0 = CurrentY
        CurrentX = x0 + 1
        CurrentY = y0 + 1
        UserControl.ForeColor = SystemColorConstants.vb3DHighlight
        UserControl.Print m_Caption
        CurrentX = x0
        CurrentY = y0
        UserControl.ForeColor = SystemColorConstants.vb3DShadow
        UserControl.Print m_Caption
        UserControl.ForeColor = oldForeColor
        
        Call Graphics.DrawBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
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
    m_hover = False
    Call delegateProperties_ctor
    Call ownProperties_Initialize
    Call delegateProperties_Initialize
End Sub

Private Sub UserControl_InitProperties()
    ownProperties_Initialize
    delegateProperties_Initialize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ownProperties_Read PropBag
    delegateProperties_Read PropBag
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ownProperties_Write PropBag
    delegateProperties_Write PropBag
End Sub

Private Sub UserControl_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    hover_Update X, Y
    If UserControl.Enabled And button = vbLeftButton Then notifyLeftButton True
    RaiseEvent MouseDown(button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    Graphics.ReleaseCapture
    hover_Update X, Y
    If UserControl.Enabled And button = vbLeftButton Then notifyLeftButton False
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
    If UserControl.Enabled Then notifyLeftButton True
    Graphics.SetCapture UserControl.hWnd
End Sub

Private Sub UserControl_GotFocus()
    If UserControl.Enabled Then updateFocus True
End Sub

Private Sub UserControl_LostFocus()
    If UserControl.Enabled Then updateFocus False
End Sub

Private Sub UserControl_Paint()
    onPaint
End Sub

