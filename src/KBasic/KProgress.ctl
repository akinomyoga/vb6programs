VERSION 5.00
Begin VB.UserControl KProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000002&
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   Tag             =   "0"
   Begin KBasic.KControlHelper Controller 
      Left            =   120
      Top             =   240
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
Attribute VB_Name = "KProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' KProgressBar
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
Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long
Dim m_BarBackColor As OLE_COLOR
Dim m_BarForeColor As OLE_COLOR

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

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_Caption As String)
    If m_Caption <> new_Caption Then
        m_Caption = new_Caption
        PropertyChanged "Caption"
        Controller.Refresh
    End If
End Property

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Long)
    If new_Value < m_Min Then new_Value = m_Min
    If new_Value > m_Max Then new_Value = m_Max
    If m_Value <> new_Value Then
        m_Value = new_Value
        PropertyChanged "Value"
        Controller.Refresh
    End If
End Property

Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Min = m_Min
End Property

Public Property Let Min(ByVal new_Min As Long)
    If new_Min > m_Max Then new_Min = m_Max
    If m_Min <> new_Min Then
        m_Min = new_Min
        PropertyChanged "Min"
        Me.Value = m_Value
        Controller.Refresh
    End If
End Property

Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Max = m_Max
End Property

Public Property Let Max(ByVal new_Max As Long)
    If new_Max < m_Min Then new_Max = m_Min
    If m_Max <> new_Max Then
        m_Max = new_Max
        PropertyChanged "Max"
        Me.Value = m_Value
        Controller.Refresh
    End If
End Property

Public Property Get BarBackColor() As OLE_COLOR
Attribute BarBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarBackColor = m_BarBackColor
End Property

Public Property Let BarBackColor(ByVal new_BarBackColor As OLE_COLOR)
    If m_BarBackColor <> new_BarBackColor Then
        m_BarBackColor = new_BarBackColor
        PropertyChanged "BarBackColor"
        Controller.Refresh
    End If
End Property

Public Property Get BarForeColor() As OLE_COLOR
Attribute BarForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarForeColor = m_BarForeColor
End Property

Public Property Let BarForeColor(ByVal new_BarForeColor As OLE_COLOR)
    If m_BarForeColor <> new_BarForeColor Then
        m_BarForeColor = new_BarForeColor
        PropertyChanged "BarForeColor"
        Controller.Refresh
    End If
End Property

Private Sub processOwnProperties(ByVal kind As PropertyOperation, PropBag As PropertyBag)
    Controller.DefineByValProperty kind, PropBag, "Caption", m_Caption, "Progress"
    Controller.DefineByValProperty kind, PropBag, "Value", m_Value, 50
    Controller.DefineByValProperty kind, PropBag, "Min", m_Min, 0
    Controller.DefineByValProperty kind, PropBag, "Max", m_Max, 100
    Controller.DefineByValProperty kind, PropBag, "BarBackColor", m_BarBackColor, RGB(0, 0, &H80)
    Controller.DefineByValProperty kind, PropBag, "BarForeColor", m_BarForeColor, ColorConstants.vbWhite
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

Private Sub doPaint()
    Dim h As Single, w As Single
    h = UserControl.ScaleHeight
    w = UserControl.ScaleWidth
    
    Dim range As Long
    range = m_Max - m_Min
    
    Dim fraction As Double
    If range = 0 Then
        fraction = 1#
    Else
        fraction = KMath.Clamp(CDbl(m_Value - m_Min) / CDbl(range), 0#, 1#)
    End If
    
    Dim v1 As Long
    v1 = CLng(w * fraction)
    
    Line (0, 0)-(v1 - 1, h - 1), Me.BarBackColor, BF
    
    Dim hDC As Long
    hDC = UserControl.hDC
    Dim hRgn_save As Long
    hRgn_save = KWin.CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(hDC, hRgn_save) = 0 Then
        KWin.DeleteObject hRgn_save
        hRgn_save = 0
    End If
    Dim save_ForeColor As OLE_COLOR
    save_ForeColor = UserControl.ForeColor
    
    Dim hRgn1 As Long
    
    hRgn1 = KWin.CreateRectRgn(0, 0, v1, h)
    If hRgn_save <> 0 Then KWin.CombineRgn hRgn1, hRgn1, hRgn_save, RGN_AND
    KWin.SelectClipRgn hDC, hRgn1
    UserControl.ForeColor = m_BarForeColor
    Controller.DrawButtonText 0, 0, w, h, kbAppearanceDefault, 0, m_Caption
    
    KWin.SetRectRgn hRgn1, v1, 0, w, h
    If hRgn_save <> 0 Then KWin.CombineRgn hRgn1, hRgn1, hRgn_save, RGN_AND
    KWin.SelectClipRgn hDC, hRgn1
    UserControl.ForeColor = save_ForeColor
    Controller.DrawButtonText 0, 0, w, h, kbAppearanceDefault, 0, m_Caption
    
    If hRgn1 <> 0 Then KWin.DeleteObject hRgn1

    UserControl.ForeColor = save_ForeColor
    KWin.SelectClipRgn hDC, hRgn_save
    If hRgn_save <> 0 Then KWin.DeleteObject hRgn_save

    KWin.DrawControlBorder Me, kbBorderControlInset, 0, 0, w, h
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub Controller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Controller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    ' m_variable = 0
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
