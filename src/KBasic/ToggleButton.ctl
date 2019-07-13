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

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Dim m_Caption As String
Dim m_Value As Boolean

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

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Boolean)
    If m_Value <> new_Value Then
        m_Value = new_Value
        Controller.Refresh
        PropertyChanged "Value"
    End If
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_Caption As String)
    If m_Caption <> new_Caption Then
        m_Caption = new_Caption
        Controller.Refresh
        PropertyChanged "Caption"
    End If
End Property

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Caption", m_Caption, "ToggleButton"
    Controller.DefineByValProperty kind, PropBag, "Value", m_Value, False
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
    Controller.SetEnabled new_Enabled
End Property

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

Public Property Get Font() As StdFont
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

Private Sub toggleState()
    If m_Value Then
        m_Value = False
    Else
        m_Value = True
    End If
    RaiseEvent Click
    Controller.Refresh
End Sub

Private Sub notifyLeftButton(ByVal state As Boolean)
    If Controller.Hover And Not state Then
        Call toggleState
    Else
        Call Controller.Refresh
    End If
End Sub

Private Sub hover_Update()
    If Controller.IsLeftPressed Then Controller.Refresh
End Sub

Private Sub doPaint()
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    text_width = UserControl.TextWidth(m_Caption)
    text_height = UserControl.TextHeight(m_Caption)
    CurrentX = (w - text_width) / 2
    CurrentY = (h - text_height) / 2
    If Controller.IsLeftPressed And Controller.Hover Then
        CurrentX = CurrentX + 1
        CurrentY = CurrentY + 1
    End If
    If Not UserControl.Enabled Then
        saveForeColor = UserControl.ForeColor
        saveX = CurrentX
        saveY = CurrentY
        UserControl.ForeColor = SystemColorConstants.vb3DHighlight
        CurrentX = saveX + 1
        CurrentY = saveY + 1
        UserControl.Print m_Caption
        UserControl.ForeColor = SystemColorConstants.vb3DShadow
        CurrentX = saveX
        CurrentY = saveY
        UserControl.Print m_Caption
        UserControl.ForeColor = saveForeColor
    Else
        UserControl.Print m_Caption
    End If

    If Controller.IsLeftPressed And Controller.Hover Then
        If Controller.HasFocus Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonInset, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderButtonInset, 0, 0, w, h)
            If Value Then UserControl.Line (4, 4)-(w - 5, h - 5), SystemColorConstants.vb3DDKShadow, B
        End If
    ElseIf Value Then
        If Controller.HasFocus Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderButtonPressed, 0, 0, w, h)
            If UserControl.Enabled Then
                UserControl.Line (4, 4)-(w - 5, h - 5), SystemColorConstants.vb3DDKShadow, B
            Else
                KWin.DrawControlBorder Me, kbBorderGroove, 4, 4, w - 3, h - 3
            End If
        End If
    Else
        If Controller.HasFocus Then
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutsetBold, 0, 0, w, h)
            Call KWin.DrawControlBorder(Me, kbBorderButtonFocus, 0, 0, w, h)
        Else
            Call KWin.DrawControlBorder(Me, kbBorderButtonOutset, 0, 0, w, h)
        End If
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub Controller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then notifyLeftButton True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseEnter(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hover_Update
End Sub

Private Sub Controller_MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hover_Update
End Sub

Private Sub Controller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then notifyLeftButton False
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

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Controller.OnMouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Controller.OnMouseMove Button, Shift, X, Y
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

