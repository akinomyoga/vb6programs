VERSION 5.00
Begin VB.UserControl ScrollBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ScrollBar.ctx":0000
   Begin KBasic.KControlHelper Controller 
      Left            =   600
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   661
      ExportsEnabled  =   -1  'True
      ExportsBackColor=   -1  'True
      ExportsForeColor=   -1  'True
      ExportsMousePointer=   -1  'True
      ExportsMouseIcon=   -1  'True
      ExportsTag      =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' ScrollBar
'' 参考 http://home.att.ne.jp/zeta/gen/excel/c04p38.htm

Public Enum KScrollAppearance
    kbScroll3D
    kbScrollFlat
    kbScrollFlat3D
End Enum

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_dragX As Single
Dim m_dragY As Single
Dim m_dragValue As Long
Dim m_button As Long
Dim m_hoverButton As Long

Private Enum HitResult
    hitNone = 0
    hitButton1 = 1
    hitButton2 = 2
    hitMargin1 = 3
    hitMargin2 = 4
    hitBar = 5
End Enum

Private Type ScrollBarGeometry
    geo_horizontal As Boolean
    geo_width As Long
    geo_height As Long
    geo_buttonSize As Long
    geo_trackSize As Long
    geo_barSize As Long
    geo_barOffset As Long
End Type

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const BAR_MIN_SIZE = 5
Const BUTTON_MIN_SIZE = 9
Const INITIAL_DELAY_FACTOR = 5

Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long
Dim m_SmallChange As Long
Dim m_Orientation As KSpinOrientation
Dim m_Delay As Long
Dim m_BarSize As Long
Dim m_BarMinSize As Long
Dim m_BarRange As Long
Dim m_LargeChange As Long
Dim m_ButtonSize As Long
Dim m_ButtonMinSize As Long
Dim m_Appearance As KScrollAppearance

Public Event Scroll()
Public Event Change()

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get Value() As Long
Attribute Value.VB_Description = "スクロールバーの現在の値を取得します"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0
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
        Controller.Refresh
    End If
End Property

Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Behavior"
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
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Behavior"
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
Attribute SmallChange.VB_ProcData.VB_Invoke_Property = ";Behavior"
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal new_SmallChange As Long)
    If m_SmallChange <> new_SmallChange Then
        m_SmallChange = new_SmallChange
        PropertyChanged "SmallChange"
    End If
End Property

Public Property Get Orientation() As KSpinOrientation
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal new_Orientation As KSpinOrientation)
    If m_Orientation <> new_Orientation Then
        m_Orientation = new_Orientation
        PropertyChanged "Orientation"
        Controller.Refresh
    End If
End Property

Public Property Get Delay() As Long
Attribute Delay.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Delay = m_Delay
End Property

Public Property Let Delay(ByVal new_Delay As Long)
    If m_Delay <> new_Delay Then
        m_Delay = new_Delay
        PropertyChanged "Delay"
    End If
End Property

Public Property Get BarSize() As Long
Attribute BarSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarSize = m_BarSize
End Property

Public Property Let BarSize(ByVal new_BarSize As Long)
    If new_BarSize <= 0 Then
        new_BarSize = -1
    ElseIf new_BarSize < m_BarMinSize Then
        new_BarSize = m_BarMinSize
    End If
    If m_BarSize <> new_BarSize Then
        m_BarSize = new_BarSize
        PropertyChanged "BarSize"
        Controller.Refresh
    End If
End Property

Public Property Get BarMinSize() As Long
Attribute BarMinSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarMinSize = m_BarMinSize
End Property

Public Property Let BarMinSize(ByVal new_BarMinSize As Long)
    If new_BarMinSize < BAR_MIN_SIZE Then new_BarMinSize = BAR_MIN_SIZE
    If m_BarMinSize <> new_BarMinSize Then
        m_BarMinSize = new_BarMinSize
        PropertyChanged "BarMinSize"
        Me.BarSize = m_BarSize
        Controller.Refresh
    End If
End Property

Public Property Get BarRange() As Long
Attribute BarRange.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarRange = m_BarRange
End Property

Public Property Let BarRange(ByVal new_BarRange As Long)
    If new_BarRange <= 0 Then new_BarRange = -1
    If m_BarRange <> new_BarRange Then
        m_BarRange = new_BarRange
        PropertyChanged "BarRange"
        Controller.Refresh
    End If
End Property

Public Property Get LargeChange() As Long
Attribute LargeChange.VB_ProcData.VB_Invoke_Property = ";Behavior"
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal new_LargeChange As Long)
    If m_LargeChange <> new_LargeChange Then
        m_LargeChange = new_LargeChange
        PropertyChanged "LargeChange"
    End If
End Property

Public Property Get ButtonSize() As Long
Attribute ButtonSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonSize = m_ButtonSize
End Property

Public Property Let ButtonSize(ByVal new_ButtonSize As Long)
    If new_ButtonSize <= 0 Then
        new_ButtonSize = -1
    ElseIf new_ButtonSize < m_ButtonMinSize Then
        new_ButtonSize = m_ButtonMinSize
    End If
    If m_ButtonSize <> new_ButtonSize Then
        m_ButtonSize = new_ButtonSize
        PropertyChanged "ButtonSize"
        Controller.Refresh
    End If
End Property

Public Property Get ButtonMinSize() As Long
Attribute ButtonMinSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonMinSize = m_ButtonMinSize
End Property

Public Property Let ButtonMinSize(ByVal new_ButtonMinSize As Long)
    If new_ButtonMinSize < BUTTON_MIN_SIZE Then new_ButtonMinSize = BUTTON_MIN_SIZE
    If m_ButtonMinSize <> new_ButtonMinSize Then
        m_ButtonMinSize = new_ButtonMinSize
        PropertyChanged "ButtonMinSize"
        Me.ButtonSize = m_ButtonSize
        Controller.Refresh
    End If
End Property

Public Property Get Appearance() As KScrollAppearance
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal new_Appearance As KScrollAppearance)
    If m_Appearance <> new_Appearance Then
        m_Appearance = new_Appearance
        PropertyChanged "Appearance"
        Controller.Refresh
    End If
End Property

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Value", m_Value, 0
    Controller.DefineByValProperty kind, PropBag, "Min", m_Min, 0
    Controller.DefineByValProperty kind, PropBag, "Max", m_Max, 10
    Controller.DefineByValProperty kind, PropBag, "SmallChange", m_SmallChange, 1
    Controller.DefineByValProperty kind, PropBag, "Orientation", m_Orientation, KSpinOrientation.kbOrientationAuto
    Controller.DefineByValProperty kind, PropBag, "Delay", m_Delay, 50
    Controller.DefineByValProperty kind, PropBag, "BarSize", m_BarSize, -1
    Controller.DefineByValProperty kind, PropBag, "BarMinSize", m_BarMinSize, BAR_MIN_SIZE
    Controller.DefineByValProperty kind, PropBag, "BarRange", m_BarRange, -1
    Controller.DefineByValProperty kind, PropBag, "LargeChange", m_LargeChange, 1
    Controller.DefineByValProperty kind, PropBag, "ButtonSize", m_ButtonSize, -1
    Controller.DefineByValProperty kind, PropBag, "ButtonMinSize", m_ButtonMinSize, BUTTON_MIN_SIZE
    Controller.DefineByValProperty kind, PropBag, "Appearance", m_Appearance, KScrollAppearance.kbScroll3D
End Sub

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    Controller.SetBackColor new_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    Controller.SetForeColor new_ForeColor
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
    If Controller.SetEnabled(new_Enabled) And Not new_Enabled Then
        m_button = 0
        m_hoverButton = 0
    End If
End Property

Public Property Get Tag() As String
    Tag = UserControl.Tag
End Property

Public Property Let Tag(ByVal new_Tag As String)
    Controller.SetTag new_Tag
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal new_MousePointer As Integer)
    Controller.SetMousePointer new_MousePointer
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByRef new_MouseIcon As IPictureDisp)
    Controller.SetMouseIcon new_MouseIcon
End Property

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Function isHorizontal() As Boolean
    Select Case m_Orientation
    Case KSpinOrientation.kbOrientationHorizontal
        isHorizontal = True
    Case KSpinOrientation.kbOrientationVertical
        isHorizontal = False
    Case Else
        isHorizontal = ScaleWidth > ScaleHeight
    End Select
End Function

Private Sub determineGeometry(ByRef geo As ScrollBarGeometry)
    geo.geo_horizontal = isHorizontal()
    If geo.geo_horizontal Then
        geo.geo_width = ScaleHeight
        geo.geo_height = ScaleWidth
    Else
        geo.geo_width = ScaleWidth
        geo.geo_height = ScaleHeight
    End If
    
    If m_ButtonSize > 0 Then
        geo.geo_buttonSize = m_ButtonSize
    Else
        geo.geo_buttonSize = KMath.MinL(geo.geo_width, geo.geo_height / 4)
    End If
    Dim minButtonSize As Long
    minButtonSize = KMath.MaxL(m_ButtonMinSize, BUTTON_MIN_SIZE)
    minButtonSize = KMath.MinL(geo.geo_height / 2, minButtonSize)
    geo.geo_buttonSize = KMath.MaxL(geo.geo_buttonSize, minButtonSize)

    geo.geo_trackSize = geo.geo_height - 2 * geo.geo_buttonSize
    If geo.geo_trackSize < BAR_MIN_SIZE Then
        geo.geo_buttonSize = geo.geo_buttonSize + CLng(geo.geo_trackSize / 2)
        geo.geo_trackSize = geo.geo_height - 2 * geo.geo_buttonSize
        geo.geo_barSize = 0
        geo.geo_barOffset = 0
        Exit Sub
    End If
    
    Dim range As Long: range = Abs(m_Max - m_Min)
    If m_BarSize > 0 Then
        geo.geo_barSize = m_BarSize
    Else
        Dim brange As Long
        If m_BarRange > 0 Then
            brange = m_BarRange
        Else
            brange = KMath.MaxL(m_LargeChange, 1)
        End If
        Dim Z As Double: Z = 1 + CDbl(range) / CDbl(brange)
        geo.geo_barSize = 2 + CLng((geo.geo_trackSize - 2) / Z)
    End If
    geo.geo_barSize = KMath.MinL(KMath.MaxL(geo.geo_barSize, m_BarMinSize), geo.geo_trackSize)
    
    Dim maxOffset As Long: maxOffset = geo.geo_trackSize - geo.geo_barSize
    Dim offset As Long: offset = Abs(m_Value - m_Min)
    geo.geo_barOffset = KMath.ClampL(CLng(maxOffset * offset / range), 0, maxOffset)
End Sub

Private Function hitTestG(ByVal X As Single, ByVal Y As Single, ByRef geo As ScrollBarGeometry) As Long
    Dim u As Single
    Dim v As Single
    If geo.geo_horizontal Then
        u = Y
        v = X
    Else
        u = X
        v = Y
    End If
    hitTestG = 0
    If 0 <= v And v < geo.geo_height And 0 <= u And u < geo.geo_width Then
        If v < geo.geo_buttonSize Then
            hitTestG = 1
        ElseIf v >= geo.geo_height - geo.geo_buttonSize Then
            hitTestG = 2
        ElseIf v < geo.geo_buttonSize + geo.geo_barOffset Then
            hitTestG = 3
        ElseIf v >= geo.geo_buttonSize + geo.geo_barOffset + geo.geo_barSize Then
            hitTestG = 4
        Else
            hitTestG = 5
        End If
    End If
End Function

Private Function hitTest(ByVal X As Single, ByVal Y As Single) As Long
    Dim geo As ScrollBarGeometry
    determineGeometry geo
    hitTest = hitTestG(X, Y, geo)
End Function

Private Sub doScroll()
    oldValue = m_Value
    isReverted = m_Min > m_Max
    If m_button = 1 Or m_button = 2 Then
        If m_button = 1 Xor isReverted Then
            m_Value = KMath.MaxL(m_Value - m_SmallChange, KMath.MinL(m_Min, m_Max))
        Else
            m_Value = KMath.MinL(m_Value + m_SmallChange, KMath.MaxL(m_Min, m_Max))
        End If
    ElseIf m_button = 3 Or m_button = 4 Then
        If m_button = 3 Xor isReverted Then
            m_Value = KMath.MaxL(m_Value - m_LargeChange, KMath.MinL(m_Min, m_Max))
        Else
            m_Value = KMath.MinL(m_Value + m_LargeChange, KMath.MaxL(m_Min, m_Max))
        End If
    End If
    If m_Value <> oldValue Then
        Controller.Refresh
        RaiseEvent Scroll
        RaiseEvent Change
    End If
End Sub

Sub leftButton_Update(ByVal state As Boolean, ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub

    oldButton = m_button
    If state Then
        m_button = hitTest(Controller.MouseX, Controller.MouseY)
        m_hoverButton = m_button
        If m_button <> 0 Then
            If m_button = 5 Then
                m_dragX = X
                m_dragY = Y
                m_dragValue = m_Value
            Else
                doScroll
                Timer1.Interval = m_Delay * INITIAL_DELAY_FACTOR
                Timer1.Enabled = True
            End If
        End If
    Else
        m_button = 0
        Timer1.Enabled = False
        Timer1.Interval = 0
    End If
    If m_button <> oldButton Then
        Controller.Refresh
    End If
End Sub

Sub OnMouseMove(ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub
    
    Dim geo As ScrollBarGeometry
    determineGeometry geo
    oldMatch = m_button = m_hoverButton
    m_hoverButton = hitTestG(X, Y, geo)
    newMatch = m_button = m_hoverButton
    If m_button = 1 Or m_button = 2 Then
        If oldMatch <> newMatch Then
            Controller.Refresh
        End If
    ElseIf m_button = 5 Then
        Dim delta As Long
        If geo.geo_horizontal Then
            delta = Controller.MouseX - m_dragX
        Else
            delta = Controller.MouseY - m_dragY
        End If
        new_Value = m_dragValue + CLng((m_Max - m_Min) * delta / (geo.geo_trackSize - geo.geo_barSize))
        min_Value = KMath.MinL(m_Min, m_Max)
        max_Value = KMath.MaxL(m_Min, m_Max)
        new_Value = KMath.ClampL(new_Value, min_Value, max_Value)
        If new_Value <> m_Value Then
            m_Value = new_Value
            Controller.Refresh
            RaiseEvent Scroll
            RaiseEvent Change
        End If
    End If
End Sub

Private Sub process_Hover(X As Single, Y As Single)
    If UserControl.Enabled And m_Appearance = kbScrollFlat3D Then
        Controller.Refresh
    End If
End Sub

Private Sub doPaint_paintButton(ByVal flags As Long, ByVal Button As Long, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    
    pressed = m_button = Button And m_button = m_hoverButton
    arrow_color = UserControl.ForeColor
    Select Case m_Appearance
    Case kbScrollFlat
        flags = flags Or kbArrowButtonFlat
        If m_button = Button Then
            If m_button = m_hoverButton Then
                Line (x1 + 1, y1 + 1)-(x2 - 2, y2 - 2), UserControl.ForeColor, BF
            Else
                Line (x1 + 1, y1 + 1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, BF
            End If
            arrow_color = UserControl.BackColor
            pressed = False
        End If
    Case kbScrollFlat3D
        If Controller.Hover Or m_button <> 0 Then
            flags = flags Or kbArrowButtonSingle
        Else
            flags = flags Or kbArrowButtonFlat
        End If
    End Select
    If pressed Then flags = flags Or kbArrowPressed
    If Not UserControl.Enabled Then flags = flags Or kbArrowDisabled

    KWin.DrawArrowButton Me, flags, x1, y1, x2, y2, arrow_color, 5, 0.6
End Sub

Private Sub doPaint_paintBar(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Select Case m_Appearance
    Case kbScrollFlat
        KWin.DrawControlBorder Me, kbBorderSinglePressed, x1, y1, x2, y2
        If m_button = 5 Then
            Line (x1 + 1, y1 + 1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, BF
        End If
    Case kbScrollFlat3D
        If Controller.Hover Or m_button <> 0 Then
            KWin.DrawControlBorder Me, kbBorderSingleOutset, x1, y1, x2, y2
        Else
            KWin.DrawControlBorder Me, kbBorderSinglePressed, x1, y1, x2, y2
        End If
    Case Else
        KWin.DrawControlBorder Me, kbBorderControlOutset, x1, y1, x2, y2
    End Select
End Sub

Private Sub doPaint_drawLine(ByRef geo As ScrollBarGeometry, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal color As OLE_COLOR)
    If geo.geo_horizontal Then
        Line (y1, x1)-(y2, x2), color
    Else
        Line (x1, y1)-(x2, y2), color
    End If
End Sub

Private Sub doPaint()
    Dim geo As ScrollBarGeometry
    determineGeometry geo
    
    Dim w As Long: w = geo.geo_width
    Dim h As Long: h = geo.geo_height
    Dim v1 As Long: v1 = geo.geo_buttonSize
    Dim v4 As Long: v4 = geo.geo_height - geo.geo_buttonSize
    If geo.geo_trackSize > 0 Then
        doPaint_drawLine geo, 0, v1, 0, v4, SystemColorConstants.vb3DShadow
        doPaint_drawLine geo, w - 1, v1, w - 1, v4, SystemColorConstants.vb3DShadow
        If Not UserControl.Enabled Then
            If geo.geo_horizontal Then
                KWin.FillChidori Me, v1, 1, v4, w - 1, SystemColorConstants.vb3DHighlight
            Else
                KWin.FillChidori Me, 1, v1, w - 1, v4, SystemColorConstants.vb3DHighlight
            End If
        Else
            Dim v2 As Long: v2 = v1 + geo.geo_barOffset
            Dim v3 As Long: v3 = v2 + geo.geo_barSize
            If geo.geo_horizontal Then
                doPaint_paintBar v2, 0, v3, w
                KWin.FillChidori Me, v1, 1, v2, w - 1, SystemColorConstants.vb3DHighlight
                KWin.FillChidori Me, v3, 1, v4, w - 1, SystemColorConstants.vb3DHighlight
            Else
                doPaint_paintBar 0, v2, w, v3
                KWin.FillChidori Me, 1, v1, w - 1, v2, SystemColorConstants.vb3DHighlight
                KWin.FillChidori Me, 1, v3, w - 1, v4, SystemColorConstants.vb3DHighlight
            End If
        End If
    End If
    If geo.geo_horizontal Then
        doPaint_paintButton kbArrowLeft, 1, 0, 0, v1, w
        doPaint_paintButton kbArrowRight, 2, v4, 0, h, w
    Else
        doPaint_paintButton kbArrowUp, 1, 0, 0, w, v1
        doPaint_paintButton kbArrowDown, 2, 0, v4, w, h
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Controller)
''
''-----------------------------------------------------------------------------

Private Sub Controller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then leftButton_Update True, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseEnter(Button As Integer, Shift As Integer, X As Single, Y As Single)
    process_Hover X, Y
End Sub

Private Sub Controller_MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
    process_Hover X, Y
End Sub

Private Sub Controller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnMouseMove X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then leftButton_Update False, X, Y
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Controller_Paint()
    doPaint
End Sub

Private Sub Controller_ProcessProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    processOwnProperties kind, PropBag
End Sub

Private Sub Timer1_Timer()
    m_hoverButton = hitTest(Controller.MouseX, Controller.MouseY)
    If m_button <> 0 And m_button = m_hoverButton Then doScroll
    Timer1.Interval = m_Delay
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    Controller.OnDblClick
End Sub

Private Sub UserControl_Initialize()
    m_button = 0
    m_hoverButton = 0
    m_dragX = 0
    m_dragY = 0
    m_dragValue = 0
    Controller.OnInitialize
End Sub

Private Sub UserControl_InitProperties()
    Controller.OnInitProperties
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Controller.OnMouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Controller.OnMouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Controller.OnMouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Paint()
    Controller.OnPaint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Controller.OnReadProperties PropBag
End Sub

Private Sub UserControl_Show()
    Controller.OnShow
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Controller.OnWriteProperties PropBag
End Sub
