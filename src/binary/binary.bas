Attribute VB_Name = "Module1"
Public unitBit As Integer
Public unitBit2 As Long

Public nowBit As Integer
Public bits(7) As Integer

Public pxWidth As Integer
Public fileopened As Boolean

Sub Main()
    pxWidth = 8
    unitBit = 2
    unitBit2 = 2 ^ unitBit
    
    nowBit = 8
    fileopened = False
    
    Call Form1.Show
End Sub

Public Function openfile(filename) As String
    If fileopened Then
        openfile = "ファイルはすでに開かれています。"
        Exit Function
    Else
        fileopened = True
    End If
    
    On Error GoTo err1
    Open filename For Binary Access Read As 1
    openfile = "ファイルは無事に開かれました。ファイルサイズ - " & LOF(1)
    nowBit = 8
    Exit Function
    
err1:
    openfile = "何らかのエラー - " & Err.number
End Function

Public Function readbits(number As Integer) As Long
    Dim k As Long
    k = 1
    readbits = 0
    For n = 1 To number
        If nowBit > 7 Then
            Call readByte
            If nowBit = -1 Then readbits = -1: Exit Function
        End If
        readbits = readbits + k * bits(nowBit)
        k = k * 2
        nowBit = nowBit + 1
    Next n
End Function

Public Sub readByte()
    If EOF(1) Then
        nowBit = -1
        Exit Sub
    End If
    Dim bytes() As Byte: bytes() = InputB(1, #1): bits(0) = bytes(0)
    For i = 0 To 6
        bits(i + 1) = Int(bits(i) / 2)
        bits(i) = bits(i) - bits(i + 1) * 2
    Next i
    nowBit = 0
End Sub

Public Function readBbyH()
    Dim bytes() As Byte
    If EOF(1) Then readBbyH = "EOF": Exit Function
    bytes() = InputB(1, 1)
    readBbyH = Hex(Int(bytes(0)))
End Function

Public Function closefile() As String
    Close #1
    fileopened = False
    closefile = "成功"
    Form1.bStop.Enabled = False
End Function
