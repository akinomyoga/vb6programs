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
      ExportsEnabled  =   -1  'True
      ExportsBackColor=   -1  'True
      ExportsForeColor=   -1  'True
      ExportsFont     =   -1  'True
      ExportsMousePointer=   -1  'True
      ExportsMouseIcon=   -1  'True
      ExportsTag      =   -1  'True
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Caption", m_Caption, "ColorButton"
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

Private Sub doPaint()
    Dim h As Single, w As Single
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    Dim pressed As Boolean, var_captionColor As OLE_COLOR, var_shiftText As Boolean
    pressed = UserControl.Enabled And Controller.IsLeftPressed And Controller.Hover
    var_captionColor = UserControl.ForeColor
    var_shiftText = pressed
    If m_Appearance = kbAppearanceFlat And pressed Then
        Line (1, 1)-(w - 2, h - 2), var_captionColor, BF
        var_captionColor = UserControl.BackColor
        var_shiftText = False
    End If
    
    Dim text_width As Single, text_height As Single
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If var_shiftText Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    
    Dim oldForeColor As OLE_COLOR
    oldForeColor = UserControl.ForeColor
    If UserControl.Enabled Then
        UserControl.ForeColor = var_captionColor
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
        ElseIf Controller.HasFocus Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutsetBold, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        End If
    End Select
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

