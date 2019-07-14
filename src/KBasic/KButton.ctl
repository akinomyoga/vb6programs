VERSION 5.00
Begin VB.UserControl KButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "KButton.ctx":0000
   Begin KBasic.KControlHelper Controller 
      Left            =   120
      Top             =   120
      _extentx        =   661
      _extenty        =   661
      exportsenabled  =   -1  'True
      exportsbackcolor=   -1  'True
      exportsforecolor=   -1  'True
      exportsfont     =   -1  'True
      exportsmousepointer=   -1  'True
      exportsmouseicon=   -1  'True
      exportstag      =   -1  'True
   End
End
Attribute VB_Name = "KButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Dim m_Caption As String
Dim m_Appearance As KControlAppearance

Public Event Click()

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

Public Property Get caption() As String
Attribute caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute caption.VB_UserMemId = 0
    caption = m_Caption
End Property

Public Property Let caption(ByVal new_Caption As String)
    If m_Caption <> new_Caption Then
        m_Caption = new_Caption
        Controller.Refresh
        PropertyChanged "Caption"
    End If
End Property

Public Property Get Appearance() As KControlAppearance
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal new_Appearance As KControlAppearance)
    If m_Appearance <> new_Appearance Then
        m_Appearance = new_Appearance
        Controller.Refresh
        PropertyChanged "Appearance"
    End If
End Property

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Caption", m_Caption, "KButton"
    Controller.DefineByValProperty kind, PropBag, "Appearance", m_Appearance, KControlAppearance.kbAppearanceDefault
End Sub

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (定義)
''
''-----------------------------------------------------------------------------

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
    Controller.SetEnabled new_Enabled
End Property

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

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef new_Font As StdFont)
    Controller.SetFont new_Font
End Property

Public Property Get Tag() As String
    Tag = UserControl.Tag
End Property

Public Property Let Tag(ByVal new_Tag As String)
    Controller.SetTag new_Tag
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal new_MousePointer As Integer)
    Controller.SetMousePointer new_MousePointer
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Behavior"
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

Private Sub updateHover()
    If UserControl.Enabled Then
        If Controller.IsLeftPressed Then
            Controller.Refresh
        ElseIf m_Appearance = kbAppearanceToolButton Or m_Appearance = kbAppearanceFlat3D Then
            Controller.Refresh
        End If
    End If
End Sub

Private Sub doMouseDown(ByVal Button As Integer)
    If UserControl.Enabled And Button = vbLeftButton Then
        Controller.Refresh
    End If
End Sub

Private Sub doMouseUp(ByVal Button As Integer)
    If UserControl.Enabled And Button = vbLeftButton Then
        Controller.Refresh
        If Controller.Hover Then RaiseEvent Click
    End If
End Sub

' DrawButtonText(x1,y1,x2,y2,color,is_pressed,is_enabled)
Private Sub doPaint_fore(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
ByVal color As OLE_COLOR, ByVal shift_text As Boolean)
    Dim text_width As Single, text_height As Single
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = x1 + Int((x2 - x1 - text_width) / 2)
    CurrentY = y1 + Int((y2 - y1 - text_height) / 2)
    If shift_text Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If

    Dim oldForeColor As OLE_COLOR
    oldForeColor = UserControl.ForeColor
    If UserControl.Enabled Then
        UserControl.ForeColor = color
        UserControl.Print m_Caption
    Else
        Dim x0 As Single, y0 As Single
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
    End If
    UserControl.ForeColor = oldForeColor
End Sub

Private Sub doPaint()
    Dim h As Single, w As Single
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    Dim Appearance As KControlAppearance
    Appearance = m_Appearance
    If Appearance = kbAppearanceDefault Then Appearance = kbAppearance3DButton
    
    Dim bflags As KButtonStateFlags
    If Not UserControl.Enabled Then
        bflags = kbButtonStateDisabled
    Else
        bflags = 0
        If Controller.IsLeftPressed And Controller.Hover Then bflags = bflags Or kbButtonStatePressed
        If Controller.IsLeftPressed Or Controller.Hover Then bflags = bflags Or kbButtonStateHovered
        If Controller.HasFocus Then bflags = bflags Or kbButtonStateFocused
    End If

    Controller.DrawButtonBackground 0, 0, w, h, Appearance, bflags
    Controller.DrawButtonText 0, 0, w, h, Appearance, bflags, m_Caption
    Controller.DrawButtonBorder 0, 0, w, h, Appearance, bflags
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub Controller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    doMouseDown Button
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseEnter(Button As Integer, Shift As Integer, X As Single, Y As Single)
    updateHover
End Sub

Private Sub Controller_MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
    updateHover
End Sub

Private Sub Controller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    doMouseUp Button
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Controller_Paint()
    doPaint
End Sub

Private Sub Controller_ProcessProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    processOwnProperties kind, PropBag
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Controller Hook)
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    Controller.OnDblClick
End Sub

Private Sub UserControl_GotFocus()
    Controller.OnGotFocus
End Sub

Private Sub UserControl_Initialize()
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

Private Sub UserControl_LostFocus()
    Controller.OnLostFocus
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

Private Sub UserControl_Resize()
    Controller.OnResize
End Sub

Private Sub UserControl_Show()
    Controller.OnShow
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Controller.OnWriteProperties PropBag
End Sub

