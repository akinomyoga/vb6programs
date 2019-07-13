VERSION 5.00
Begin VB.UserControl SpinButton 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "SpinButton.ctx":0000
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
Attribute VB_Name = "SpinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' SpinButton
'' 参考 http://home.att.ne.jp/zeta/gen/excel/c04p38.htm

Public Enum KSpinOrientation
    kbOrientationAuto = -1
    kbOrientationVertical = 0
    kbOrientationHorizontal = 1
End Enum

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_button As Long
Dim m_hoverButton As Long

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const INITIAL_DELAY_FACTOR = 5

Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long
Dim m_SmallChange As Long
Dim m_Orientation As KSpinOrientation
Dim m_Delay As Long

Public Event SpinUp()
Public Event SpinDown()
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

Public Property Get Orientation() As KSpinOrientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal new_Orientation As KSpinOrientation)
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

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Value", m_Value, 0
    Controller.DefineByValProperty kind, PropBag, "Min", m_Min, 0
    Controller.DefineByValProperty kind, PropBag, "Max", m_Max, 10
    Controller.DefineByValProperty kind, PropBag, "SmallChange", m_SmallChange, 1
    Controller.DefineByValProperty kind, PropBag, "Orientation", m_Orientation, KSpinOrientation.kbOrientationAuto
    Controller.DefineByValProperty kind, PropBag, "Delay", m_Delay, 50
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
    Controller.SetBackColor new_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    Controller.SetForeColor new_ForeColor
End Property

Public Property Get Enabled() As Boolean
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
        isHorizontal = Width > Height
    End Select
End Function

Function hitTest(ByVal X As Single, ByVal Y As Single) As Long
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

Private Sub doSpin()
    oldValue = m_Value
    isReverted = m_Min > m_Max
    If m_button = 1 Xor isHorizontal() Xor isReverted Then
        m_Value = KMath.MinL(m_Value + m_SmallChange, KMath.MaxL(m_Min, m_Max))
    Else
        m_Value = KMath.MaxL(m_Value - m_SmallChange, KMath.MinL(m_Min, m_Max))
    End If
    If m_Value <> oldValue Then
        If m_Value > oldValue Then
            RaiseEvent SpinUp
        ElseIf m_Value < oldValue Then
            RaiseEvent SpinDown
        End If
        RaiseEvent Change
    End If
End Sub

Sub leftButton_Update(ByVal state As Boolean, ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub

    oldButton = m_button
    If state Then
        m_button = hitTest(X, Y)
        m_hoverButton = m_button
        If m_button <> 0 Then doSpin
        Timer1.Interval = m_Delay * INITIAL_DELAY_FACTOR
        Timer1.Enabled = True
    Else
        m_button = 0
        Timer1.Enabled = False
        Timer1.Interval = 0
    End If
    If m_button <> oldButton Then
        Controller.Refresh
    End If
End Sub

Sub doMouseMove(ByVal X As Long, ByVal Y As Long)
    If Not UserControl.Enabled Then Exit Sub
    If m_button <> 0 Then
        oldMatch = m_button = m_hoverButton
        m_hoverButton = hitTest(X, Y)
        newMatch = m_button = m_hoverButton
        If oldMatch <> newMatch Then
            Controller.Refresh
        End If
    End If
End Sub

Sub doPaint_paintButton2(ByVal flags As Long, ByVal Button As Long, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    
    flags = flags Or kbArrowButtonInset
    pressed = m_button = Button And m_button = m_hoverButton
    If pressed Then flags = flags Or kbArrowPressed
    If Not UserControl.Enabled Then flags = flags Or kbArrowDisabled
    KWin.DrawArrowButton Me, flags, x1, y1, x2, y2, UserControl.ForeColor, 5, 1#
End Sub

Sub doPaint()
    w = KMath.FloorL(ScaleWidth, 2)
    h = KMath.FloorL(ScaleHeight, 2)
    If isHorizontal() Then
        m = Int(w / 2)
        doPaint_paintButton2 kbArrowLeft, 1, 0, 0, m, h
        doPaint_paintButton2 kbArrowRight, 2, m, 0, w, h
    Else
        m = Int(h / 2)
        doPaint_paintButton2 kbArrowUp, 1, 0, 0, w, m
        doPaint_paintButton2 kbArrowDown, 2, 0, m, w, h
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub Controller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        leftButton_Update True, X, Y
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    doMouseMove X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        leftButton_Update False, X, Y
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Controller_Paint()
    doPaint
End Sub

Private Sub Controller_ProcessProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    processOwnProperties kind, PropBag
End Sub

Private Sub Timer1_Timer()
    If m_button <> 0 And m_button = m_hoverButton Then doSpin
    Timer1.Interval = m_Delay
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Controller Hook)
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    Controller.OnDblClick
End Sub

Private Sub UserControl_Initialize()
    m_button = 0
    m_hoverButton = 0
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
