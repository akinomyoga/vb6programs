VERSION 5.00
Begin VB.UserControl ScrollBar 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' ScrollBar
'' 参考 http://home.att.ne.jp/zeta/gen/excel/c04p38.htm

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_leftButton As Long
Dim m_mouseX As Single
Dim m_mouseY As Single
Dim m_button As Long
Dim m_hoverButton As Long

Dim m_barSize As Long
Dim m_barRange As Long
Dim m_barMinSize As Long
Dim m_largeChange As Long

Const BAR_MIN_SIZE = 5

Private Enum HitResult
    hitNone = 0
    hitButton1 = 1
    hitButton2 = 2
    hitMargin1 = 3
    hitMargin2 = 4
    hitBar = 5
End Enum

Private Type ScrollBarGeometry
    m_horizontal As Boolean
    m_width As Long
    m_height As Long
    m_buttonSize As Long
    m_trackSize As Long
    m_barSize As Long
    m_barOffset As Long
End Type

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_Value = 0
Const default_Min = 0
Const default_Max = 10
Const default_SmallChange = 1
Const default_Orientation = -1
Const default_Delay = 1000

Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long
Dim m_SmallChange As Long
Dim m_Orientation As SpinButtonOrientation
Dim m_Delay As Long

Public Event SpinUp()
Public Event SpinDown()
Public Event Change()

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_BackColor = SystemColorConstants.vbButtonFace
Const default_ForeColor = SystemColorConstants.vbButtonText
Const default_Enabled = True
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

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Long)
    min_Value = KMath.MinL(m_Min, m_Max)
    max_Value = KMath.MaxL(m_Min, m_Max)
    new_Value = KMath.ClampL(new_Value, min_Value, max_Value)
    If m_Value <> new_Value Then
        m_Value = new_Value
        PropertyChanged "Value"
        RaiseEvent Change
    End If
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal new_Min As Long)
    If m_Min <> new_Min Then
        m_Min = new_Min
        PropertyChanged "Min"
        Me.Value = m_Value
    End If
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal new_Max As Long)
    If m_Max <> new_Max Then
        m_Max = new_Max
        PropertyChanged "Max"
        Me.Value = m_Value
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal new_SmallChange As Long)
    If m_SmallChange <> new_SmallChange Then
        m_SmallChange = new_SmallChange
        PropertyChanged "SmallChange"
    End If
End Property

Public Property Get Orientation() As SpinButtonOrientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal new_Orientation As SpinButtonOrientation)
    If m_Orientation <> new_Orientation Then
        m_Orientation = new_Orientation
        PropertyChanged "Orientation"
    End If
End Property

Public Property Get Delay() As Long
    Delay = m_Delay
End Property

Public Property Let Delay(ByVal new_Delay As Long)
    If m_Delay <> new_Delay Then
        m_Delay = new_Delay
        PropertyChanged "Delay"
    End If
End Property

Sub ownProperties_Initialize()
    m_Value = default_Value
    m_Min = 0
    m_Max = default_Max
    m_SmallChange = default_SmallChange
    m_Orientation = default_Orientation
    m_Delay = default_Delay
End Sub

Sub ownProperties_Read(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", default_Value)
    m_Min = PropBag.ReadProperty("Min", default_Min)
    m_Max = PropBag.ReadProperty("Max", default_Max)
    m_SmallChange = PropBag.ReadProperty("SmallChange", default_SmallChange)
    m_Orientation = PropBag.ReadProperty("Orientation", default_Orientation)
    m_Delay = PropBag.ReadProperty("Delay", default_Delay)
End Sub

Sub ownProperties_Write(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, default_Value)
    Call PropBag.WriteProperty("Min", m_Min, default_Min)
    Call PropBag.WriteProperty("Max", m_Max, default_Max)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, default_SmallChange)
    Call PropBag.WriteProperty("Orientation", m_Orientation, default_Orientation)
    Call PropBag.WriteProperty("Delay", m_Delay, default_Delay)
End Sub

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    If UserControl.BackColor <> new_BackColor Then
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

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
    If UserControl.Enabled <> new_Enabled Then
        UserControl.Enabled = new_Enabled
        If Not new_Enabled Then
            m_leftButton = False
            m_button = 0
            m_hoverButton = 0
        End If
        UserControl.Refresh
        PropertyChanged "Enabled"
    End If
End Property

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
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    UserControl.Enabled = default_Enabled
    UserControl.Tag = default_Tag
    UserControl.MousePointer = default_MousePointer
    Set UserControl.MouseIcon = default_MouseIcon
End Sub

Sub delegateProperties_Read(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", default_Enabled)
    UserControl.Tag = PropBag.ReadProperty("Tag", default_Tag)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", default_MousePointer)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", default_MouseIcon)
End Sub

Sub delegateProperties_Write(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, default_BackColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, default_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, default_Enabled)
    Call PropBag.WriteProperty("Tag", UserControl.Tag, default_Tag)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, default_MousePointer)
    Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, default_MouseIcon)
End Sub

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Function isHorizontal() As Boolean
    Select Case m_Orientation
    Case SpinButtonOrientation.kbOrientationHorizontal
        isHorizontal = True
    Case SpinButtonOrientation.kbOrientationVertical
        isHorizontal = False
    Case Else
        isHorizontal = ScaleWidth > ScaleHeight
    End Select
End Function

Private Sub determineGeometry(ByRef geo As ScrollBarGeometry)
    geo.m_horizontal = isHorizontal()
    If geo.m_horizontal Then
        geo.m_width = ScaleHeight
        geo.m_height = ScaleWidth
    Else
        geo.m_width = ScaleWidth
        geo.m_height = ScaleHeight
    End If
    
    Dim minButtonSize As Long: minButtonSize = KMath.MinL(geo.m_height / 2, 9)
    geo.m_buttonSize = KMath.MaxL(minButtonSize, KMath.MinL(geo.m_width, geo.m_height / 4))

    geo.m_trackSize = geo.m_height - 2 * geo.m_buttonSize
    If geo.m_trackSize < BAR_MIN_SIZE Then
        geo.m_buttonSize = geo.m_buttonSize + CLng(geo.m_trackSize / 2)
        geo.m_trackSize = geo.m_height - 2 * geo.m_buttonSize
        geo.m_barSize = 0
        geo.m_barOffset = 0
        Exit Sub
    End If
    
    Dim range As Long: range = Abs(m_Max - m_Min)
    If m_barSize > 0 Then
        geo.m_barSize = m_barSize
    Else
        Dim brange As Long
        If m_barRange > 0 Then
            brange = m_barRange
        Else
            brange = KMath.MaxL(m_largeChange, 1)
        End If
        Dim Z As Double: Z = 1 + CDbl(range) / CDbl(brange)
        geo.m_barSize = 2 + CLng((geo.m_trackSize - 2) / Z)
    End If
    geo.m_barSize = KMath.MinL(KMath.MaxL(geo.m_barSize, m_barMinSize), geo.m_trackSize)
    
    Dim maxOffset As Long: maxOffset = geo.m_trackSize - geo.m_barSize
    Dim offset As Long: offset = Abs(m_Value - m_Min)
    geo.m_barOffset = KMath.ClampL(CLng(maxOffset * offset / range), 0, maxOffset)
End Sub

Function hitTest(ByVal X As Single, ByVal Y As Single) As Long
    Dim geo As ScrollBarGeometry
    determineGeometry geo
    
    Dim u As Single
    Dim v As Single
    If geo.m_horizontal Then
        u = Y
        v = X
    Else
        u = X
        v = Y
    End If
    histTest = 0
    If 0 <= v And v < geo.m_height And 0 <= u And u < geo.m_width Then
        If v < geo.m_buttonSize Then
            hitTest = 1
        ElseIf v >= geo.m_height - geo.m_buttonSize Then
            hitTest = 2
        ElseIf v < geo.m_buttonSize + geo.m_barOffset Then
            hitTest = 3
        ElseIf v >= geo.m_buttonSize + geo.m_barOffset + geo.m_barSize Then
            hitTest = 4
        Else
            hitTest = 5
        End If
    End If
End Function

Sub leftButton_Update(ByVal state As Boolean, ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub
    m_mouseX = X
    m_mouseY = Y

    If m_leftButton = state Then Exit Sub
    m_leftButton = state
    oldButton = m_button
    If state Then
        m_button = hitTest(X, Y)
        m_hoverButton = m_button
        If m_button <> 0 Then
            oldValue = m_Value
            isReverted = m_Min > m_Max
            If m_button = 1 Or m_button = 2 Then
                If m_button = 1 Xor isReverted Then
                    m_Value = KMath.MaxL(m_Value - m_SmallChange, KMath.MinL(m_Min, m_Max))
                Else
                    m_Value = KMath.MinL(m_Value + m_SmallChange, KMath.MaxL(m_Min, m_Max))
                End If
            End If
            If m_Value <> oldValue Then
                If m_Value > oldValue Then
                    RaiseEvent SpinUp
                ElseIf m_Value < oldValue Then
                    RaiseEvent SpinDown
                End If
                RaiseEvent Change
            End If
        End If
    Else
        m_button = 0
    End If
    If m_button <> oldButton Then
        UserControl.Refresh
    End If
End Sub

Sub onMouseMove(ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub
    m_mouseX = X
    m_mouseY = Y
    If m_button <> 0 Then
        oldMatch = m_button = m_hoverButton
        m_hoverButton = hitTest(X, Y)
        newMatch = m_button = m_hoverButton
        If oldMatch <> newMatch Then
            UserControl.Refresh
        End If
    End If
End Sub

Sub onPaint_paintButton(ByVal flags As Long, ByVal button As Long, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    
    pressed = m_button = button And m_button = m_hoverButton
    If pressed Then flags = flags Or kbArrowPressed
    If Not UserControl.Enabled Then flags = flags Or kbArrowDisabled
    KWin.DrawArrowButton Me, flags, x1, y1, x2, y2, UserControl.ForeColor, 5, 0.6
End Sub

Private Sub onPaint_drawLine(ByRef geo As ScrollBarGeometry, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal color As OLE_COLOR)
    If geo.m_horizontal Then
        Line (y1, x1)-(y2, x2), color
    Else
        Line (x1, y1)-(x2, y2), color
    End If
End Sub

Private Sub onPaint()
    Dim geo As ScrollBarGeometry
    determineGeometry geo
    
    Dim w As Long: w = geo.m_width
    Dim h As Long: h = geo.m_height
    Dim v1 As Long: v1 = geo.m_buttonSize
    Dim v4 As Long: v4 = geo.m_height - geo.m_buttonSize
    If geo.m_trackSize > 0 Then
        Dim v2 As Long: v2 = v1 + geo.m_barOffset
        Dim v3 As Long: v3 = v2 + geo.m_barSize
        onPaint_drawLine geo, 0, v1, 0, v4, SystemColorConstants.vb3DShadow
        onPaint_drawLine geo, w - 1, v1, w - 1, v4, SystemColorConstants.vb3DShadow
        If geo.m_horizontal Then
            KWin.DrawBorder Me, kbBorderControlOutset, v2, 0, v3, w
            KWin.FillChidori Me, v1, 1, v2, w - 1, SystemColorConstants.vb3DHighlight
            KWin.FillChidori Me, v3, 1, v4, w - 1, SystemColorConstants.vb3DHighlight
        Else
            KWin.DrawBorder Me, kbBorderControlOutset, 0, v2, w, v3
            KWin.FillChidori Me, 1, v1, w - 1, v2, SystemColorConstants.vb3DHighlight
            KWin.FillChidori Me, 1, v3, w - 1, v4, SystemColorConstants.vb3DHighlight
        End If
    End If
    If geo.m_horizontal Then
        onPaint_paintButton kbArrowLeft, 1, 0, 0, v1, w
        onPaint_paintButton kbArrowRight, 2, v4, 0, h, w
    Else
        onPaint_paintButton kbArrowUp, 1, 0, 0, w, v1
        onPaint_paintButton kbArrowDown, 2, 0, v4, w, h
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    leftButton_Update True, m_mouseX, m_mouseY
    KWin.SetCapture UserControl.hWnd
End Sub

Private Sub UserControl_Initialize()
    m_leftButton = False
    m_mouseX = 0
    m_mouseY = 0
    m_button = 0
    m_hoverButton = 0
    Call delegateProperties_ctor
    Call ownProperties_Initialize
    Call delegateProperties_Initialize
End Sub

Private Sub UserControl_InitProperties()
    Call ownProperties_Initialize
    Call delegateProperties_Initialize
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

Private Sub UserControl_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    leftButton_Update True, X, Y
    RaiseEvent MouseDown(button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    onMouseMove X, Y
    RaiseEvent MouseMove(button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    KWin.ReleaseCapture
    leftButton_Update False, X, Y
    RaiseEvent MouseUp(button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    onPaint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ownProperties_Read PropBag
    delegateProperties_Read PropBag
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ownProperties_Write PropBag
    delegateProperties_Write PropBag
End Sub


