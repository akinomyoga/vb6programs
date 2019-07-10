VERSION 5.00
Begin VB.UserControl SpinButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "SpinButton.ctx":0000
End
Attribute VB_Name = "SpinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' SpinButton
'' 参考 http://home.att.ne.jp/zeta/gen/excel/c04p38.htm

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_leftButton As Integer
Dim m_mouseX As Single
Dim m_mouseY As Single
Dim m_button As Integer

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Public Enum SpinButtonOrientation
    kbOrientationAuto = -1
    kbOrientationVertical = 0
    kbOrientationHorizontal = 1
End Enum

Const default_Value = 0
Const default_Min = 0
Const default_Max = 10
Const default_SmallChange = 1
Const default_Orientation = -1
Const default_Delay = 1000

Dim m_Value As Integer
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_SmallChange As Integer
Dim m_Orientation As SpinButtonOrientation
Dim m_Delay As Integer

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

Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Integer)
    If m_Value <> new_Value Then
        m_Value = new_Value
        PropertyChanged "Value"
        RaiseEvent Change
    End If
End Property

Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal new_Min As Integer)
    If m_Min <> new_Min Then
        m_Min = new_Min
        PropertyChanged "Min"
    End If
End Property

Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal new_Max As Integer)
    If m_Max <> new_Max Then
        m_Max = new_Max
        PropertyChanged "Max"
    End If
End Property

Public Property Get SmallChange() As Integer
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal new_SmallChange As Integer)
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

Public Property Get Delay() As Integer
    Delay = m_Delay
End Property

Public Property Let Delay(ByVal new_Delay As Integer)
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
        isHorizontal = Width > Height
    End Select
End Function

Function hitTest(ByVal X As Single, ByVal Y As Integer) As Integer
    Dim pos1 As Single, max1 As Single
    Dim pos2 As Single, max2 As Single
    If isHorizontal() Then
        pos1 = X
        max1 = ScaleWidth
        pos2 = Y
        max2 = ScaleHeight
    Else
        pos1 = Y
        max1 = ScaleHeight
        pos2 = X
        max2 = ScaleWidth
    End If
    histTest = 0
    If 0 <= pos1 And pos1 < max1 And 0 <= pos2 And pos2 < max2 Then
        If pos1 < Int(max1 / 2) Then
            hitTest = 1
        Else
            hitTest = 2
        End If
    End If
End Function

Sub leftButton_Update(ByVal state As Boolean, ByVal X As Integer, ByVal Y As Integer)
    m_mouseX = X
    m_mouseY = Y
    If m_leftButton = state Then Exit Sub
    m_leftButton = state
    oldButton = m_button
    If state Then
        m_button = hitTest(X, Y)
        If m_button <> 0 Then
            oldValue = m_Value
            If m_button = 1 Xor isHorizontal() Then
                m_Value = KMath.MinI(m_Value + m_SmallChange, m_Max)
            Else
                m_Value = KMath.MaxI(m_Value - m_SmallChange, m_Min)
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

Sub onMouseMove(ByVal X As Integer, ByVal Y As Integer)
    m_mouseX = X
    m_mouseY = Y
    If m_button <> 0 Then
        If m_button <> hitTest(X, Y) Then
            m_button = 0
            UserControl.Refresh
        End If
    End If
End Sub

Sub onPaint_PaintArrow(ByVal direction As Integer, ByVal pressed As Boolean, _
    ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)

    Dim w As Integer: w = x2 - x1
    Dim h As Integer: h = y2 - y1
    Dim x0 As Integer: x0 = x1 + w / 2 - 1
    Dim y0 As Integer: y0 = y1 + h / 2 - 1
    If pressed Then
        x0 = x0 + 1
        y0 = y0 + 1
    End If

    Dim aw As Integer
    Dim vx As Integer, vy As Integer
    Dim ux As Integer, uy As Integer
    Select Case direction
    Case 1 't
        aw = KMath.MinI(h - 5, (w - 7) / 2)
        vx = 0: vy = 1: ux = 1: uy = 0
    Case 2 'b
        aw = KMath.MinI(h - 5, (w - 7) / 2)
        vx = 0: vy = -1: ux = 1: uy = 0
    Case 3 'l
        aw = KMath.MinI(w - 5, (h - 7) / 2)
        vx = 1: vy = 0: ux = 0: uy = 1
    Case 4 'r
        aw = KMath.MinI(w - 5, (h - 7) / 2)
        vx = -1: vy = 0: ux = 0: uy = 1
    End Select
    aw = KMath.ClampI(aw, 2, 5)
    
    x0 = x0 - vx * Int(aw / 2)
    y0 = y0 - vy * Int(aw / 2)
    For i = 0 To aw - 1
        px = x0 + i * vx
        py = y0 + i * vy
        PSet (px, py), ForeColor
        For j = 1 To i
            PSet (px + j * ux, py + j * uy), ForeColor
            PSet (px - j * ux, py - j * uy), ForeColor
        Next j
    Next i
End Sub

Sub onPaint_PaintButton(ByVal direction As Integer, ByVal pressed As Boolean, _
    ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)

    onPaint_PaintArrow direction, pressed, x1, y1, x2, y2
    If pressed Then
        kbb = kbBorderButtonInset
    Else
        kbb = kbBorderButtonOutset
    End If
    Graphics.DrawBorder Me, kbb, x1, y1, x2, y2
End Sub

Sub onPaint()
    w = KMath.CeilI(ScaleWidth, 2)
    h = KMath.CeilI(ScaleHeight, 2)
    If isHorizontal() Then
        m = Int(w / 2)
        onPaint_PaintButton 3, m_button = 1, 0, 0, m, h
        onPaint_PaintButton 4, m_button = 2, m, 0, w, h
    Else
        m = Int(h / 2)
        onPaint_PaintButton 1, m_button = 1, 0, 0, w, m
        onPaint_PaintButton 2, m_button = 2, 0, m, w, h
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    leftButton_Update True, m_mouseX, m_mouseY
End Sub

Private Sub UserControl_Initialize()
    m_leftButton = False
    m_mouseX = 0
    m_mouseY = 0
    m_button = 0
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    leftButton_Update True, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    onMouseMove X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    leftButton_Update False, X, Y
    RaiseEvent MouseUp(Button, Shift, X, Y)
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
