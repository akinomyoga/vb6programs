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
   Begin KBasic.KControlHelper Controller 
      Left            =   120
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   661
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Enum KControlAppearance
    kbAppearanceDefault
    kbAppearance3D
    kbAppearance3DInset
    kbAppearance3DSingle
    kbAppearance3DButton
    kbAppearanceFlat
    kbAppearanceFlat3D
    kbAppearanceToolButton
    kbAppearanceGroove
End Enum

''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim m_hasFocus As Boolean

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const default_Caption = "ColorButton"
Const default_Appearance = KControlAppearance.kbAppearanceDefault

Dim m_Caption As String
Dim m_Appearance As KControlAppearance

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

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = 0
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_Caption As String)
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

Private Sub ownProperties_Initialize()
    m_Caption = default_Caption
    m_Appearance = default_Appearance
End Sub

Private Sub ownProperties_Read(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", default_Caption)
    m_Appearance = PropBag.ReadProperty("Appearance", default_Appearance)
End Sub

Private Sub ownProperties_Write(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, default_Caption)
    Call PropBag.WriteProperty("Appearance", m_Appearance, default_Appearance)
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
    If UserControl.Enabled <> new_Enabled Then
        UserControl.Enabled = new_Enabled
        If Not new_Enabled Then
            m_hasFocus = False
        End If
        Controller.Refresh
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    If UserControl.BackColor <> new_BackColor Then
        UserControl.BackColor = new_BackColor
        Controller.Refresh
        PropertyChanged "BackColor"
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
    If UserControl.ForeColor <> new_ForeColor Then
        UserControl.ForeColor = new_ForeColor
        Controller.Refresh
        PropertyChanged "ForeColor"
    End If
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef new_Font As StdFont)
    If UserControl.Font <> new_Font Then
        Set UserControl.Font = new_Font
        Controller.Refresh
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

Private Sub delegateProperties_ctor()
    Set default_MouseIcon = Nothing
End Sub

Private Sub delegateProperties_Initialize()
    UserControl.Enabled = default_Enabled
    UserControl.BackColor = default_BackColor
    UserControl.ForeColor = default_ForeColor
    If default_Font Is Nothing Then Set default_Font = UserControl.Font
    UserControl.Tag = default_Tag
    UserControl.MousePointer = default_MousePointer
    Set UserControl.MouseIcon = default_MouseIcon
End Sub

Private Sub delegateProperties_Read(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", default_Enabled)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", default_BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", default_ForeColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", getDefaultFont())
    UserControl.Tag = PropBag.ReadProperty("Tag", default_Tag)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", default_MousePointer)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", default_MouseIcon)
End Sub

Private Sub delegateProperties_Write(PropBag As PropertyBag)
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

Private Sub updateFocus(ByVal state As Boolean)
    If m_hasFocus <> state Then
        m_hasFocus = state
        Controller.Refresh
    End If
End Sub

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

Private Sub doPaint()
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    pressed = UserControl.Enabled And Controller.IsLeftPressed And Controller.Hover
    var_captionColor = UserControl.ForeColor
    var_shiftText = pressed
    If m_Appearance = kbAppearanceFlat And pressed Then
        Line (1, 1)-(w - 2, h - 2), var_captionColor, BF
        var_captionColor = UserControl.BackColor
        var_shiftText = False
    End If
    
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If var_shiftText Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    
    oldForeColor = UserControl.ForeColor
    If UserControl.Enabled Then
        UserControl.ForeColor = var_captionColor
        UserControl.Print m_Caption
    Else
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
        
    Select Case m_Appearance
    Case kbAppearance3D
        If pressed Then
            Call KWin.DrawControlBorder(Me, kbBorderSinglePressed, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderControlOutset, 0, 0, w, h)
        End If
    Case kbAppearance3DInset
        If pressed Then
            Call KWin.DrawControlBorder(Me, kbBorderControlInset, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderControlOutset, 0, 0, w, h)
        End If
    Case kbAppearance3DSingle
        If pressed Then
            Call KWin.DrawControlBorder(Me, kbBorderSingleInset, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderSingleOutset, 0, 0, w, h)
        End If
    Case kbAppearanceGroove
        If pressed Then
            Call KWin.DrawControlBorder(Me, kbBorderControlInset, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderGroove, 0, 0, w, h)
        End If
    Case kbAppearanceFlat
        KWin.DrawControlBorder Me, kbBorderSinglePressed, 0, 0, w, h
    Case kbAppearanceFlat3D
        If UserControl.Enabled Then
            If pressed Then
                KWin.DrawControlBorder Me, kbBorderSingleInset, 0, 0, w, h
            ElseIf Controller.IsLeftPressed Or Controller.Hover Then
                KWin.DrawControlBorder Me, kbBorderSingleOutset, 0, 0, w, h
            Else
                KWin.DrawControlBorder Me, kbBorderSinglePressed, 0, 0, w, h
            End If
        Else
            KWin.DrawControlBorder Me, kbBorderSinglePressed, 0, 0, w, h
        End If
    Case kbAppearanceToolButton
        If UserControl.Enabled Then
            If pressed Then
                KWin.DrawControlBorder Me, kbBorderSingleInset, 0, 0, w, h
            ElseIf Controller.IsLeftPressed Or Controller.Hover Then
                KWin.DrawControlBorder Me, kbBorderSingleOutset, 0, 0, w, h
            End If
        End If
    Case Else ' kbAppearance3DButton
        If Not UserControl.Enabled Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        ElseIf pressed Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        ElseIf m_hasFocus Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutsetBold, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        End If
    End Select
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Controller)
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

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    Controller.OnDblClick
End Sub

Private Sub UserControl_Initialize()
    m_hasFocus = False
    Call delegateProperties_ctor
    Call ownProperties_Initialize
    Call delegateProperties_Initialize
End Sub

Private Sub UserControl_InitProperties()
    ownProperties_Initialize
    delegateProperties_Initialize
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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ownProperties_Read PropBag
    delegateProperties_Read PropBag
End Sub

Private Sub UserControl_Show()
    Controller.OnShow
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ownProperties_Write PropBag
    delegateProperties_Write PropBag
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

Private Sub UserControl_GotFocus()
    If UserControl.Enabled Then updateFocus True
End Sub

Private Sub UserControl_LostFocus()
    If UserControl.Enabled Then updateFocus False
End Sub

Private Sub UserControl_Paint()
    Controller.OnPaint
End Sub

