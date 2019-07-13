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


''-----------------------------------------------------------------------------
''
'' 内部変数
''
''-----------------------------------------------------------------------------

Dim user As UserControl

Dim m_mouseButton As Integer
Dim m_mouseShift As Integer
Dim m_mouseX As Single
Dim m_mouseY As Single

Dim m_button As Integer
Dim m_hover As Boolean

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

''-----------------------------------------------------------------------------
''
'' 委譲プロパティ (宣言)
''
''-----------------------------------------------------------------------------

Const fixed_Width = 375
Const fixed_Height = 375

''-----------------------------------------------------------------------------
''
'' 独自プロパティ (実装)
''
''-----------------------------------------------------------------------------

Public Property Get MouseButton() As Integer
Attribute MouseButton.VB_MemberFlags = "400"
    MouseButton = m_button
End Property

Public Property Get MouseX() As Integer
Attribute MouseX.VB_MemberFlags = "400"
    MouseX = m_mouseX
End Property

Public Property Get MouseY() As Integer
Attribute MouseY.VB_MemberFlags = "400"
    MouseY = m_mouseY
End Property

Public Property Get Hover() As Integer
Attribute Hover.VB_MemberFlags = "400"
    Hover = m_hover
End Property

''-----------------------------------------------------------------------------
''
'' 処理
''
''-----------------------------------------------------------------------------

Private Sub initializeUserControl()
    Set user = KWin.GetUserControl(UserControl.Parent)
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

Private Function hitTest(ByVal X As Single, ByVal Y As Single) As Boolean
    hitTest = 0 <= X And X < user.ScaleWidth And 0 <= Y And Y < user.ScaleHeight
End Function

Private Sub processMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button And Not (m_button And Button) Then
        m_button = m_button Or Button
        RaiseEvent MouseDown(Button, m_mouseShift, m_mouseX, m_mouseY)
    End If
    safeMouseCapture
End Sub

Private Sub processMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_mouseX = X And m_mouseY = Y Then Exit Sub

    new_hover = hitTest(X, Y)
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

Private Sub processMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (m_button And Button) <> 0 Then
        RaiseEvent MouseUp(m_button, m_mouseShift, m_mouseX, m_mouseY)
        m_button = m_button And Not Button
    End If
    If hitTest(X, Y) Then
        safeMouseCapture ' VB6 が勝手に Release してしまう様なので
    ElseIf m_button = 0 Then
        safeMouseRelease
    End If
End Sub

Public Sub Refresh()
    If user.AutoRedraw Then
        user.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), user.BackColor, BF
        RaiseEvent Paint
    Else
        user.Refresh
    End If
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録 (Parent)
''
''-----------------------------------------------------------------------------
' マウスイベントは MouseDown, MouseUp, Click / DblClick, MouseUp の順で発生するそうだ。
' http://cya.sakura.ne.jp/vb/MSHFlexGrid_Event.htm

Public Sub OnDblClick()
    initializeUserControl
    processMouseDown MouseButtonConstants.vbLeftButton, m_mouseShift, m_mouseX, m_mouseY
End Sub

Public Sub OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    initializeUserControl
    processMouseMove Button, Shift, X, Y
    processMouseDown Button, Shift, X, Y
End Sub

Public Sub OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    initializeUserControl
    processMouseMove Button, Shift, X, Y
End Sub

Public Sub OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    initializeUserControl
    processMouseMove Button, Shift, X, Y
    processMouseUp Button, Shift, X, Y
End Sub

Public Sub OnShow()
    initializeUserControl
    If user.AutoRedraw Then Refresh
    Set user = Nothing ' 何故かこれがないとクラッシュする
End Sub

Public Sub OnPaint()
    initializeUserControl
    RaiseEvent Paint
End Sub

''-----------------------------------------------------------------------------
''
'' イベント登録
''
''-----------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    Set user = Nothing
    m_button = 0
    m_hover = False
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Width", fixed_Width, fixed_Width)
    Call PropBag.WriteProperty("Height", fixed_Height, fixed_Height)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = fixed_Width
    UserControl.Height = fixed_Height
End Sub
