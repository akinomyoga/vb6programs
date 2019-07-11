Attribute VB_Name = "KWin"

Public Enum KWinBorderStyle
    kbBorderControlInset
    kbBorderControlOutset
    kbBorderButtonOutset
    kbBorderButtonPressed
    kbBorderButtonInset
    kbBorderButtonOutsetBold
    kbBorderButtonInsetBold
    kbBorderButtonFocus
    kbBorderSingleOutset
    kbBorderSinglePressed
    kbBorderSingleInset
    kbBorderGroove
    kbBorderRidge
End Enum

Public Enum KWinArrowButtonFlags
    kbArrowUp = 0
    kbArrowDown = 1
    kbArrowRight = 2
    kbArrowLeft = 3
    kbArrowDirectionMask = 3
    
    kbArrowDisabled = 4
    kbArrowPressed = 8
    kbArrowInset = 16
End Enum


Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' https://stackoverflow.com/questions/863039/problems-passing-in-a-usercontrol-as-a-parameter-in-vb6
Public Function GetUserControl(ByRef oObj As Object) As UserControl
    Dim pControl As UserControl
    Call CopyMemory(pControl, ObjPtr(oObj), 4)
    Set GetUserControl = pControl
    Call CopyMemory(pControl, 0&, 4)
End Function

Private Sub drawBorder_double(ByRef user As UserControl, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
    ByVal lt1 As OLE_COLOR, ByVal lt2 As OLE_COLOR, ByVal rb2 As OLE_COLOR, ByVal rb1 As OLE_COLOR)
    user.Line (x1, y1)-(x2 - 1, y1), lt1
    user.Line (x1, y1)-(x1, y2 - 1), lt1
    user.Line (x1 + 1, y1 + 1)-(x2 - 1, y1 + 1), lt2
    user.Line (x1 + 1, y1 + 1)-(x1 + 1, y2 - 1), lt2
    user.Line (x1 + 1, y2 - 2)-(x2 - 1, y2 - 2), rb2
    user.Line (x2 - 2, y1 + 1)-(x2 - 2, y2 - 1), rb2
    user.Line (x1, y2 - 1)-(x2, y2 - 1), rb1
    user.Line (x2 - 1, y1)-(x2 - 1, y2), rb1
End Sub

Private Sub drawBorder_single(ByRef user As UserControl, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
    ByVal lt1 As OLE_COLOR, ByVal rb1 As OLE_COLOR)
    user.Line (x1, y1)-(x2 - 1, y1), lt1
    user.Line (x1, y1)-(x1, y2 - 1), lt1
    user.Line (x1, y2 - 1)-(x2, y2 - 1), rb1
    user.Line (x2 - 1, y1)-(x2 - 1, y2), rb1
End Sub

Private Sub drawBorder_impl(ByRef user As UserControl, ByVal style As KWinBorderStyle, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    Select Case style
    Case KWinBorderStyle.kbBorderSingleOutset
        drawBorder_single user, x1, y1, x2, y2, SystemColorConstants.vb3DHighlight, SystemColorConstants.vb3DDKShadow
    Case KWinBorderStyle.kbBorderSingleInset
        drawBorder_single user, x1, y1, x2, y2, SystemColorConstants.vb3DDKShadow, SystemColorConstants.vb3DHighlight
    Case KWinBorderStyle.kbBorderSinglePressed
        drawBorder_single user, x1, y1, x2, y2, SystemColorConstants.vb3DShadow, SystemColorConstants.vb3DShadow
    Case KWinBorderStyle.kbBorderButtonOutset
        drawBorder_double user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DHighlight, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow
    Case KWinBorderStyle.kbBorderButtonInset
        drawBorder_double user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DDKShadow, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DHighlight
    Case KWinBorderStyle.kbBorderControlOutset
        drawBorder_double user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DHighlight, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow
    Case KWinBorderStyle.kbBorderControlInset
        drawBorder_double user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DHighlight
    Case KWinBorderStyle.kbBorderButtonPressed
        fs = user.FillStyle
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        user.Line (x1 + 1, y1 + 1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, B
        user.FillStyle = fs
    Case KWinBorderStyle.kbBorderButtonInsetBold
        fs = user.FillStyle
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        user.FillStyle = fs
        Call drawBorder_impl(user, kbBorderButtonInset, x1 + 1, y1 + 1, x2 - 1, y2 - 1)
    Case KWinBorderStyle.kbBorderButtonOutsetBold
        fs = user.FillStyle
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        user.FillStyle = fs
        Call drawBorder_impl(user, kbBorderButtonOutset, x1 + 1, y1 + 1, x2 - 1, y2 - 1)
    Case KWinBorderStyle.kbBorderButtonFocus
        For X = x1 + 5 To x2 - 5 Step 2
            user.PSet (X, y1 + 4), user.ForeColor
            user.PSet (X, y2 - 5), user.ForeColor
        Next X
        For Y = y1 + 5 To y2 - 5 Step 2
            user.PSet (x1 + 4, Y), user.ForeColor
            user.PSet (x2 - 5, Y), user.ForeColor
        Next Y
    Case KWinBorderStyle.kbBorderGroove
        fs = user.FillStyle
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DHighlight, B
        user.Line (x1, y1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, B
        user.FillStyle = fs
    Case KWinBorderStyle.kbBorderRidge
        fs = user.FillStyle
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DShadow, B
        user.Line (x1, y1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DHighlight, B
        user.FillStyle = fs
    End Select
End Sub

Public Sub DrawBorder(ByRef ctrl As Object, ByVal style As KWinBorderStyle, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    Dim user As UserControl
    Set user = GetUserControl(ctrl)
    Call drawBorder_impl(user, style, x1, y1, x2, y2)
End Sub

Private Sub drawArrowButton_arrow(ByRef user As UserControl, ByVal flags As KWinArrowButtonFlags, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal color As OLE_COLOR, ByVal maxArrowSize As Long, ByVal maxArrowRate As Double)

    Dim w As Long: w = x2 - x1
    Dim h As Long: h = y2 - y1
    Dim x0 As Long: x0 = x1 + w / 2 - 1
    Dim y0 As Long: y0 = y1 + h / 2 - 1
    pressed = (flags And KWinArrowButtonFlags.kbArrowPressed) <> 0
    If pressed Then
        x0 = x0 + 1
        y0 = y0 + 1
    End If

    Dim aw As Long
    Dim vx As Long, vy As Long, vm As Long
    Dim ux As Long, uy As Long, um As Long
    Select Case flags And KWinArrowButtonFlags.kbArrowDirectionMask
    Case kbArrowUp
        vm = h: um = w
        vx = 0: vy = 1: ux = 1: uy = 0
    Case kbArrowDown
        vm = h: um = w
        vx = 0: vy = -1: ux = 1: uy = 0
    Case kbArrowLeft
        vm = w: um = h
        vx = 1: vy = 0: ux = 0: uy = 1
    Case kbArrowRight
        vm = w: um = h
        vx = -1: vy = 0: ux = 0: uy = 1
    End Select
    aw = KMath.MinL(vm - 5, KMath.MinL((um - 7) / 2, CLng((um - 4) * maxArrowRate / 2)))
    aw = KMath.ClampL(aw, 2, maxArrowSize)
    
    x0 = x0 - vx * Int(aw / 2)
    y0 = y0 - vy * Int(aw / 2)
    For i = 0 To aw - 1
        px = x0 + i * vx
        py = y0 + i * vy
        user.PSet (px, py), color
        For j = 1 To i
            user.PSet (px + j * ux, py + j * uy), color
            user.PSet (px - j * ux, py - j * uy), color
        Next j
    Next i
End Sub

Private Sub drawArrowButton_impl(ByRef user As UserControl, ByVal flags As KWinArrowButtonFlags, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal color As OLE_COLOR, ByVal maxArrowSize As Long, ByVal maxArrowRate As Double)

    If (flags And KWinArrowButtonFlags.kbArrowDisabled) <> 0 Then
        drawArrowButton_arrow user, flags Or kbArrowPressed, x1, y1, x2, y2, SystemColorConstants.vb3DHighlight, maxArrowSize, maxArrowRate
        drawArrowButton_arrow user, flags And Not kbArrowPressed, x1, y1, x2, y2, SystemColorConstants.vb3DShadow, maxArrowSize, maxArrowRate
    Else
        drawArrowButton_arrow user, flags, x1, y1, x2, y2, color, maxArrowSize, maxArrowRate
    End If
    
    Dim kbb As KWinBorderStyle
    If (flags And KWinArrowButtonFlags.kbArrowPressed) <> 0 Then
        If (flags And KWinArrowButtonFlags.kbArrowInset) <> 0 Then
            kbb = KWinBorderStyle.kbBorderButtonInset
        Else
            kbb = KWinBorderStyle.kbBorderSinglePressed
        End If
    Else
        If (flags And KWinArrowButtonFlags.kbArrowInset) <> 0 Then
            kbb = KWinBorderStyle.kbBorderButtonOutset
        Else
            kbb = KWinBorderStyle.kbBorderControlOutset
        End If
    End If
    KWin.drawBorder_impl user, kbb, x1, y1, x2, y2
End Sub

Public Sub DrawArrowButton(ByRef ctrl As Object, ByVal flags As KWinArrowButtonFlags, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
    ByVal color As OLE_COLOR, ByVal maxArrowSize As Long, ByVal maxArrowRate As Double)
    Dim user As UserControl
    Set user = GetUserControl(ctrl)
    drawArrowButton_impl user, flags, x1, y1, x2, y2, color, maxArrowSize, maxArrowRate
End Sub

Private Sub FillChidori_impl(ByRef user As UserControl, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As OLE_COLOR)
    If x1 >= x2 Then Exit Sub
    If y1 >= y2 Then Exit Sub
    For X = x1 To x2 Step 2
        line_length = KMath.MinL(x2 - X, y2 - y1)
        user.Line (X, y1)-(X + line_length, y1 + line_length), color
    Next X
    For Y = y1 + 2 To y2 Step 2
        line_length = KMath.MinL(x2 - x1, y2 - Y)
        user.Line (x1, Y)-(x1 + line_length, Y + line_length), color
    Next Y
End Sub

Public Sub FillChidori(ByRef ctrl As Object, _
    ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As OLE_COLOR)
    Dim user As UserControl
    Set user = GetUserControl(ctrl)
    FillChidori_impl user, x1, y1, x2, y2, color
End Sub


