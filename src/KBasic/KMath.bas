Attribute VB_Name = "KMath"
Public Function Min(ByVal X As Double, ByVal Y As Double) As Double
    If X <= Y Then
        Min = X
    Else
        Min = Y
    End If
End Function

Public Function MinI(ByVal X As Integer, ByVal Y As Integer) As Integer
    If X <= Y Then
        MinI = X
    Else
        MinI = Y
    End If
End Function

Public Function MinL(ByVal X As Long, ByVal Y As Long) As Long
    If X <= Y Then
        MinL = X
    Else
        MinL = Y
    End If
End Function

Public Function Max(ByVal X As Double, ByVal Y As Double) As Double
    If X >= Y Then
        Max = X
    Else
        Max = Y
    End If
End Function

Public Function MaxI(ByVal X As Integer, ByVal Y As Integer) As Integer
    If X >= Y Then
        MaxI = X
    Else
        MaxI = Y
    End If
End Function

Public Function MaxL(ByVal X As Long, ByVal Y As Long) As Long
    If X >= Y Then
        MaxL = X
    Else
        MaxL = Y
    End If
End Function

Public Function Clamp(ByVal Value As Double, ByVal Min As Double, ByVal Max As Double) As Double
    If Value < Min Then
        Clamp = Min
    ElseIf Value > Max Then
        Clamp = Max
    Else
        Clamp = Value
    End If
End Function

Public Function ClampI(ByVal Value As Integer, ByVal Min As Integer, ByVal Max As Integer) As Integer
    If Value < Min Then
        ClampI = Min
    ElseIf Value > Max Then
        ClampI = Max
    Else
        ClampI = Value
    End If
End Function

Public Function ClampL(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    If Value < Min Then
        ClampL = Min
    ElseIf Value > Max Then
        ClampL = Max
    Else
        ClampL = Value
    End If
End Function

Public Function FloorI(ByVal Value As Integer, ByVal Modulo As Integer) As Integer
    If Value > 0 Then
        FloorI = Value - Value Mod Modulo
    Else
        FloorI = Value + (-Value) Mod Modulo
    End If
End Function

Public Function FloorL(ByVal Value As Long, ByVal Modulo As Long) As Long
    If Value > 0 Then
        FloorL = Value - Value Mod Modulo
    Else
        FloorL = Value + (-Value) Mod Modulo
    End If
End Function

