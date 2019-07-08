Attribute VB_Name = "Graphics"

Public Enum BorderStyle
    ControlInset
    ControlOutset
    ButtonOutset
    ButtonPressed
    ButtonInset
    ButtonOutsetBold
    ButtonInsetBold
    ButtonFocus
    SingleOutset
    SingleInset
    Groove
    Ridge
End Enum

' https://stackoverflow.com/questions/863039/problems-passing-in-a-usercontrol-as-a-parameter-in-vb6
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function GetUserControl(ByRef oObj As Object) As UserControl
    Dim pControl As UserControl
    Call CopyMemory(pControl, ObjPtr(oObj), 4)
    Set GetUserControl = pControl
    Call CopyMemory(pControl, 0&, 4)
End Function

Sub DrawBorderDouble(ByRef user As UserControl, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
    ByVal lt1 As OLE_COLOR, ByVal lt2 As OLE_COLOR, ByVal rb2 As OLE_COLOR, ByVal rb1 As OLE_COLOR)
    user.Line (x1, y1)-(x2 - 1, y1), lt1
    user.Line (x1, y1)-(x1, y2 - 1), lt1
    user.Line (x1 + 1, y1 + 1)-(x2 - 1, x1 + 1), lt2
    user.Line (x1 + 1, y1 + 1)-(x1 + 1, y2 - 1), lt2
    user.Line (x1 + 1, y2 - 2)-(x2 - 1, y2 - 2), rb2
    user.Line (x2 - 2, y1 + 1)-(x2 - 2, y2 - 1), rb2
    user.Line (x1, y2 - 1)-(x2, y2 - 1), rb1
    user.Line (x2 - 1, y1)-(x2 - 1, y2), rb1
End Sub

Sub DrawBorderImpl(ByRef user As UserControl, ByVal style As BorderStyle, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    Select Case style
    Case BorderStyle.SingleOutset
        user.Line (x1, y1)-(x2 - 1, y1), SystemColorConstants.vb3DHighlight
        user.Line (x1, y1)-(x1, y2 - 1), SystemColorConstants.vb3DHighlight
        user.Line (x1, y2 - 1)-(x2, y2 - 1), SystemColorConstants.vb3DDKShadow
        user.Line (x2 - 1, y1)-(x2 - 1, y2), SystemColorConstants.vb3DDKShadow
    Case BorderStyle.SingleInset
        user.Line (x1, y1)-(x2 - 1, y1), SystemColorConstants.vb3DDKShadow
        user.Line (x1, y1)-(x1, y2 - 1), SystemColorConstants.vb3DDKShadow
        user.Line (x1, y2 - 1)-(x2, y2 - 1), SystemColorConstants.vb3DHighlight
        user.Line (x2 - 1, y1)-(x2 - 1, y2), SystemColorConstants.vb3DHighlight
    Case BorderStyle.ButtonOutset
        DrawBorderDouble user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DHighlight, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow
    Case BorderStyle.ButtonInset
        DrawBorderDouble user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow, _
            SystemColorConstants.vb3DHighlight, _
            SystemColorConstants.vb3DLight
    Case BorderStyle.ControlOutset
        DrawBorderDouble user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DHighlight, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow
    Case BorderStyle.ControlInset
        DrawBorderDouble user, x1, y1, x2, y2, _
            SystemColorConstants.vb3DShadow, _
            SystemColorConstants.vb3DDKShadow, _
            SystemColorConstants.vb3DLight, _
            SystemColorConstants.vb3DHighlight
    Case BorderStyle.ButtonPressed
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        user.Line (x1 + 1, y1 + 1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, B
    Case BorderStyle.ButtonInsetBold
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        Call DrawBorderImpl(user, ButtonInset, x1 + 1, y1 + 1, x2 - 1, y2 - 1)
    Case BorderStyle.ButtonOutsetBold
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1, y1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DDKShadow, B
        Call DrawBorderImpl(user, ButtonOutset, x1 + 1, y1 + 1, x2 - 1, y2 - 1)
    Case BorderStyle.ButtonFocus
        For x = x1 + 5 To x2 - 5 Step 2
            user.PSet (x, y1 + 4), user.ForeColor
            user.PSet (x, y2 - 5), user.ForeColor
        Next x
        For y = y1 + 5 To y2 - 5 Step 2
            user.PSet (x1 + 4, y), user.ForeColor
            user.PSet (x2 - 5, y), user.ForeColor
        Next y
    Case BorderStyle.Groove
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DHighlight, B
        user.Line (x1, y1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DShadow, B
    Case BorderStyle.Ridge
        user.FillStyle = FillStyleConstants.vbFSTransparent
        user.Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), SystemColorConstants.vb3DShadow, B
        user.Line (x1, y1)-(x2 - 2, y2 - 2), SystemColorConstants.vb3DHighlight, B
    End Select
End Sub

Public Sub DrawBorder(ByRef ctrl As Object, ByVal style As BorderStyle, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    Dim user As UserControl
    Set user = GetUserControl(ctrl)
    fs = user.FillStyle
    ds = user.DrawStyle
    Call DrawBorderImpl(user, style, x1, y1, x2, y2)
    user.FillStyle = fs
    user.DrawStyle = ds
End Sub
