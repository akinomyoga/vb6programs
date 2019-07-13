VERSION 5.00
Begin VB.UserControl KControlHelper 
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   InvisibleAtRuntime=   -1  'True
   Picture         =   "KControlHelper.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   375
End
Attribute VB_Name = "KControlHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PropertyOperation
    kbPropertyInit
    kbPropertyRead
    kbPropertyWrite
End Enum

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Const fixed_Width = 375
Const fixed_Height = 375

Dim user As UserControl
Dim m_userDepth As Integer

Dim m_mouseButton As Integer
Dim m_mouseShift As Integer
Dim m_mouseX As Single
Dim m_mouseY As Single

Dim m_button As Integer
Dim m_hover As Boolean
Dim m_hasFocus As Boolean

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Paint()
Public Event ProcessProperties(ByVal kind As PropertyOperation, ByRef PropBag As PropertyBag)

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Dim m_exportsEnabled As Boolean
Dim m_exportsBackColor As Boolean
Dim m_exportsForeColor As Boolean
Dim m_exportsFont As Boolean
Dim m_exportsTag As Boolean
Dim m_exportsMousePointer As Boolean
Dim m_exportsMouseIcon As Boolean

Const default_Enabled = True
Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Dim default_Font As StdFont
Const default_Tag = ""
Const default_MousePointer = MousePointerConstants.vbDefault
Dim default_MouseIcon As IPictureDisp

''-----------------------------------------------------------------------------
''
'' Utility
''
''-----------------------------------------------------------------------------

Public Sub DefineByValProperty(ByVal kind As PropertyOperation, ByRef PropBag As PropertyBag, _
ByVal Name As String, ByRef Variable, defaultValue)
    Select Case kind
    Case kbPropertyInit
        Variable = defaultValue
    Case kbPropertyRead
        Variable = PropBag.ReadProperty(Name, defaultValue)
    Case kbPropertyWrite
        PropBag.WriteProperty Name, Variable, defaultValue
    End Select
End Sub

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (実装)
''
''-----------------------------------------------------------------------------

Public Property Get MouseButton() As Integer
Attribute MouseButton.VB_MemberFlags = "400"
    MouseButton = m_button
End Property

Public Property Get MouseX() As Single
Attribute MouseX.VB_MemberFlags = "400"
    MouseX = m_mouseX
End Property

Public Property Get MouseY() As Single
Attribute MouseY.VB_MemberFlags = "400"
    MouseY = m_mouseY
End Property

Public Property Get Hover() As Boolean
Attribute Hover.VB_MemberFlags = "400"
    Hover = m_hover
End Property

Public Property Get HasFocus() As Boolean
    HasFocus = m_hasFocus
End Property

Public Property Get IsLeftPressed() As Boolean
Attribute IsLeftPressed.VB_MemberFlags = "400"
    IsLeftPressed = (m_button And MouseButtonConstants.vbLeftButton) <> 0
End Property

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (設定)
''
''-----------------------------------------------------------------------------

Public Property Get ExportsEnabled() As Boolean
    ExportsEnabled = m_exportsEnabled
End Property

Public Property Let ExportsEnabled(ByVal newValue As Boolean)
    m_exportsEnabled = newValue
End Property

Public Property Get ExportsBackColor() As Boolean
    ExportsBackColor = m_exportsBackColor
End Property

Public Property Let ExportsBackColor(ByVal newValue As Boolean)
    m_exportsBackColor = newValue
End Property

Public Property Get ExportsForeColor() As Boolean
    ExportsForeColor = m_exportsForeColor
End Property

Public Property Let ExportsForeColor(ByVal newValue As Boolean)
    m_exportsForeColor = newValue
End Property

Public Property Get ExportsFont() As Boolean
    ExportsFont = m_exportsFont
End Property

Public Property Let ExportsFont(ByVal newValue As Boolean)
    m_exportsFont = newValue
End Property

Public Property Get ExportsTag() As Boolean
    ExportsTag = m_exportsTag
End Property

Public Property Let ExportsTag(ByVal newValue As Boolean)
    m_exportsTag = newValue
End Property

Public Property Get ExportsMousePointer() As Boolean
    ExportsMousePointer = m_exportsMousePointer
End Property

Public Property Let ExportsMousePointer(ByVal newValue As Boolean)
    m_exportsMousePointer = newValue
End Property

Public Property Get ExportsMouseIcon() As Boolean
    ExportsMouseIcon = m_exportsMouseIcon
End Property

Public Property Let ExportsMouseIcon(ByVal newValue As Boolean)
    m_exportsMouseIcon = newValue
End Property

Private Sub processOwnProperties(ByVal kind As PropertyOperation, Optional ByRef PropBag As PropertyBag = Nothing)
    DefineByValProperty kind, PropBag, "ExportsEnabled", m_exportsEnabled, False
    DefineByValProperty kind, PropBag, "ExportsBackColor", m_exportsBackColor, False
    DefineByValProperty kind, PropBag, "ExportsForeColor", m_exportsForeColor, False
    DefineByValProperty kind, PropBag, "ExportsFont", m_exportsFont, False
    DefineByValProperty kind, PropBag, "ExportsMousePointer", m_exportsMousePointer, False
    DefineByValProperty kind, PropBag, "ExportsMouseIcon", m_exportsMouseIcon, False
    DefineByValProperty kind, PropBag, "ExportsTag", m_exportsTag, False
End Sub

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (委譲)
''
''-----------------------------------------------------------------------------

Public Function SetEnabled(ByVal new_Enabled As Boolean, Optional ByVal toRefresh = True) As Boolean
    incrementUserControl
    SetEnabled = user.Enabled <> new_Enabled
    If SetEnabled Then
        user.Enabled = new_Enabled
        user.PropertyChanged "Enabled"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetBackColor(ByVal new_BackColor As OLE_COLOR, Optional ByVal toRefresh = True) As Boolean
    incrementUserControl
    SetBackColor = user.BackColor <> new_BackColor
    If SetBackColor Then
        user.BackColor = new_BackColor
        user.PropertyChanged "BackColor"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetForeColor(ByVal new_ForeColor As OLE_COLOR, Optional ByVal toRefresh = True) As Boolean
    incrementUserControl
    SetForeColor = user.ForeColor <> new_ForeColor
    If SetForeColor Then
        user.ForeColor = new_ForeColor
        user.PropertyChanged "ForeColor"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetFont(ByRef new_Font As StdFont, Optional ByVal toRefresh = True) As Boolean
    incrementUserControl
    SetFont = user.Font <> new_Font
    If SetFont Then
        Set user.Font = new_Font
        user.PropertyChanged "Font"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetTag(ByVal new_Tag As String, Optional ByVal toRefresh = False) As Boolean
    incrementUserControl
    SetTag = user.Tag <> new_Tag
    If SetTag Then
        user.Tag = new_Tag
        user.PropertyChanged "Tag"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetMousePointer(ByVal new_MousePointer As Integer, Optional ByVal toRefresh = False) As Boolean
    incrementUserControl
    SetMousePointer = user.MousePointer <> new_MousePointer
    If SetMousePointer Then
        user.MousePointer = new_MousePointer
        user.PropertyChanged "MousePointer"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Public Function SetMouseIcon(ByRef new_MouseIcon As IPictureDisp, Optional ByVal toRefresh = False) As Boolean
    incrementUserControl
    SetMouseIcon = user.MouseIcon <> new_MouseIcon
    If SetMouseIcon Then
        Set user.MouseIcon = new_MouseIcon
        user.PropertyChanged "MouseIcon"
        If toRefresh Then Me.Refresh
    End If
    decrementUserControl
End Function

Private Function getDefaultFont() As StdFont
    If default_Font Is Nothing Then
        Set getDefaultFont = Ambient.Font
    Else
        Set getDefaultFont = default_Font
    End If
End Function

Private Sub delegateProperties_ctor()
    Set default_MouseIcon = Nothing
End Sub

Private Sub delegateProperties_Init()
    If m_exportsEnabled Then user.Enabled = default_Enabled
    If m_exportsBackColor Then user.BackColor = default_BackColor
    If m_exportsForeColor Then user.ForeColor = default_ForeColor
    If m_exportsFont Then If default_Font Is Nothing Then Set default_Font = user.Font
    If m_exportsTag Then user.Tag = default_Tag
    If m_exportsMousePointer Then user.MousePointer = default_MousePointer
    If m_exportsMouseIcon Then Set user.MouseIcon = default_MouseIcon
End Sub

Private Sub delegateProperties_Read(PropBag As PropertyBag)
    If m_exportsEnabled Then user.Enabled = PropBag.ReadProperty("Enabled", default_Enabled)
    If m_exportsBackColor Then user.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
    If m_exportsForeColor Then user.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
    If m_exportsFont Then Set user.Font = PropBag.ReadProperty("Font", getDefaultFont())
    If m_exportsTag Then user.Tag = PropBag.ReadProperty("Tag", default_Tag)
    If m_exportsMousePointer Then user.MousePointer = PropBag.ReadProperty("MousePointer", default_MousePointer)
    If m_exportsMouseIcon Then Set user.MouseIcon = PropBag.ReadProperty("MouseIcon", default_MouseIcon)
End Sub

Private Sub delegateProperties_Write(PropBag As PropertyBag)
    If m_exportsEnabled Then PropBag.WriteProperty "Enabled", user.Enabled, default_Enabled
    If m_exportsBackColor Then PropBag.WriteProperty "BackColor", user.BackColor, default_BackColor
    If m_exportsForeColor Then PropBag.WriteProperty "ForeColor", user.ForeColor, default_ForeColor
    If m_exportsFont Then PropBag.WriteProperty "Font", user.Font, getDefaultFont()
    If m_exportsTag Then PropBag.WriteProperty "Tag", user.Tag, default_Tag
    If m_exportsMousePointer Then PropBag.WriteProperty "MousePointer", user.MousePointer, default_MousePointer
    If m_exportsMouseIcon Then PropBag.WriteProperty "MouseIcon", user.MouseIcon, default_MouseIcon
End Sub

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Private Sub incrementUserControl()
    If m_userDepth = 0 Then Set user = KWin.GetUserControl(UserControl.Parent)
    m_userDepth = m_userDepth + 1
End Sub

Private Sub decrementUserControl()
    m_userDepth = m_userDepth - 1
    If m_userDepth = 0 Then Set user = Nothing ' 何故かこれがないとクラッシュする
End Sub

Private Sub safeMouseCapture()
    If KWin.GetCapture() = 0 Then
        KWin.SetCapture user.hWnd
    End If
End Sub

Private Sub safeMouseRelease()
    If KWin.GetCapture() = user.hWnd Then
        KWin.ReleaseCapture
    End If
End Sub

Private Sub updateFocus(ByVal state As Boolean)
    If m_hasFocus <> state Then
        m_hasFocus = state
        Me.Refresh
    End If
End Sub

Private Function hitTest(ByVal X As Single, ByVal Y As Single) As Boolean
    hitTest = 0 <= X And X < user.ScaleWidth And 0 <= Y And Y < user.ScaleHeight
End Function

Private Sub doMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button And Not (m_button And Button) Then
        m_button = m_button Or Button
        RaiseEvent MouseDown(Button, m_mouseShift, m_mouseX, m_mouseY)
    End If
    safeMouseCapture
End Sub

Private Sub doMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_mouseX = X And m_mouseY = Y Then Exit Sub

    Dim new_hover As Boolean: new_hover = hitTest(X, Y)
    If new_hover <> m_hover Then
        m_hover = new_hover
        If new_hover Then
            RaiseEvent MouseEnter(Button, Shift, X, Y)
            safeMouseCapture
        Else
            If m_button = 0 Then safeMouseRelease
            RaiseEvent MouseLeave(m_mouseButton, Shift, m_mouseX, m_mouseY)
        End If
    End If

    m_mouseButton = Button
    m_mouseShift = Shift
    m_mouseX = X
    m_mouseY = Y
    RaiseEvent MouseMove(m_button, m_mouseShift, m_mouseX, m_mouseY)
End Sub

Private Sub doMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (m_button And Button) <> 0 Then
        m_button = m_button And Not Button
        RaiseEvent MouseUp(Button, m_mouseShift, m_mouseX, m_mouseY)
    End If
    If hitTest(X, Y) Then
        safeMouseCapture ' VB6 が勝手に Release してしまう様なので
    ElseIf m_button = 0 Then
        safeMouseRelease
    End If
End Sub

Public Sub Refresh()
    incrementUserControl
    If user.AutoRedraw Then
        user.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), user.BackColor, BF
        RaiseEvent Paint
    Else
        user.Refresh
    End If
    decrementUserControl
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Parent)
''
''-----------------------------------------------------------------------------
' マウスイベントは MouseDown, MouseUp, Click / DblClick, MouseUp の順で発生するそうだ。
' http://cya.sakura.ne.jp/vb/MSHFlexGrid_Event.htm

Public Sub OnDblClick()
    incrementUserControl
    doMouseDown MouseButtonConstants.vbLeftButton, m_mouseShift, m_mouseX, m_mouseY
    decrementUserControl
End Sub

Public Sub OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    incrementUserControl
    doMouseMove Button, Shift, X, Y
    doMouseDown Button, Shift, X, Y
    decrementUserControl
End Sub

Public Sub OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    incrementUserControl
    doMouseMove Button, Shift, X, Y
    decrementUserControl
End Sub

Public Sub OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    incrementUserControl
    doMouseMove Button, Shift, X, Y
    doMouseUp Button, Shift, X, Y
    decrementUserControl
End Sub

Public Sub OnShow()
    incrementUserControl
    If user.AutoRedraw Then Refresh
    decrementUserControl
End Sub

Public Sub OnResize()
    incrementUserControl
    If user.AutoRedraw Then Refresh
    decrementUserControl
End Sub

Public Sub OnPaint()
    incrementUserControl
    RaiseEvent Paint
    decrementUserControl
End Sub

Public Sub OnInitialize()
    incrementUserControl
    delegateProperties_ctor
    delegateProperties_Init
    RaiseEvent ProcessProperties(kbPropertyInit, Nothing)
    decrementUserControl
End Sub

Public Sub OnInitProperties()
    incrementUserControl
    delegateProperties_Init
    RaiseEvent ProcessProperties(kbPropertyInit, Nothing)
    decrementUserControl
End Sub

Public Sub OnReadProperties(ByRef PropBag As PropertyBag)
    incrementUserControl
    delegateProperties_Read PropBag
    RaiseEvent ProcessProperties(kbPropertyRead, PropBag)
    decrementUserControl
End Sub

Public Sub OnWriteProperties(ByRef PropBag As PropertyBag)
    incrementUserControl
    delegateProperties_Write PropBag
    RaiseEvent ProcessProperties(kbPropertyWrite, PropBag)
    decrementUserControl
End Sub

Public Sub OnGotFocus()
    updateFocus True
End Sub

Public Sub OnLostFocus()
    updateFocus False
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    Set user = Nothing
    m_mouseButton = 0
    m_mouseShift = 0
    m_mouseX = 0
    m_mouseY = 0
    m_button = 0
    m_hover = False
    m_hasFocus = False
    UserControl_InitProperties
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
    processOwnProperties kbPropertyInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    processOwnProperties kbPropertyRead, PropBag
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    processOwnProperties kbPropertyWrite, PropBag
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
End Sub
