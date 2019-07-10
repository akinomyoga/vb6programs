Attribute VB_Name = "KMath"
Public Function Min(X As Double, Y As Double) As Double
    If X <= Y Then
        Min = X
    Else
        Min = Y
    End If
End Function

Public Function MinI(X As Integer, Y As Integer) As Integer
    If X <= Y Then
        MinI = X
    Else
        MinI = Y
    End If
End Function

Public Function Max(X As Double, Y As Double) As Double
    If X >= Y Then
        Max = X
    Else
        Max = Y
    End If
End Function

Public Function MaxI(X As Integer, Y As Integer) As Integer
    If X >= Y Then
        MaxI = X
    Else
        MaxI = Y
    End If
End Function

Public Function ClampI(Value As Integer, Min As Integer, Max As Integer) As Integer
    If Value < Min Then
        ClampI = Min
    ElseIf Value > Max Then
        ClampI = Max
    Else
        ClampI = Value
    End If
End Function

Public Function CeilI(Value As Integer, Modulo As Integer) As Integer
    If Value > 0 Then
        CeilI = Value - Value Mod Modulo
    Else
        CeilI = Value + (-Value) Mod Modulo
    End If
End Function

