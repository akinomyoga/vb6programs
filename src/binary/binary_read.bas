Attribute VB_Name = "Module2"
Public Sub readImage()
    Dim x As Integer, y As Integer, lx As Integer
    Dim bits0 As Long
    x = 0
    y = 0
    lx = 0
    Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 8 / unitBit / pxWidth / 320 + 1) * 15 * (pxWidth + 1)
    If Form1.chkInv.Value = 1 Then
        Do
            bits0 = readbits(unitBit)
            If bits0 = -1 Then Exit Do
            Call draw(bits0, x + lx, y)
            x = x - 1
            If x = -1 Then
                x = pxWidth - 1
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
    Else
        Do
            bits0 = readbits(unitBit)
            If bits0 = -1 Then Exit Do
            Call draw(bits0, x + lx, y)
            x = x + 1
            If x = pxWidth Then
                x = 0
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
    End If
End Sub

Public Sub draw(num As Long, x As Integer, y As Integer)
    If num < 0 Or unitBit2 <= num Then
        Picture1.PSet (x, y), RGB(255, 0, 0)
        Exit Sub
    End If
    
    Select Case unitBit
    Case 8
        g = Int(num / 8)
        b = num - g * 8
        r = Int(g / 4)
        g = g - r * 4
        color1 = RGB(36 * r, 85 * g, 36 * b)
    Case 24
        g = Int(num / 256)
        b = num - g * 256
        r = Int(g / 256)
        g = g - r * 256
        color1 = RGB(r, g, b)
    Case Else
        color1 = RGB(0, 0, 0)
    End Select
    Form1.Picture1.PSet (x, y), color1
End Sub


'#####################################################

Public Sub readImage1bt()
    Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 8 / pxWidth / 320 + 1) * 15 * (pxWidth + 1)
    Dim x As Integer, y As Integer, lx As Integer
    Dim bits0 As Integer
    y = 0
    lx = 0
    
    If Form1.chkInv.Value = 1 Then
        x = pxWidth - 1
        Do
            bits0 = readbits(1)
            If bits0 = -1 Then Exit Do
            Call draw1bt(bits0, x + lx, y)
            x = x - 1
            If x = -1 Then
                x = pxWidth - 1
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1 '9
            End If
        Loop
        Exit Sub
    Else
        x = 0
        Do
            bits0 = readbits(1)
            If bits0 = -1 Then Exit Do
            Call draw1bt(bits0, x + lx, y)
            x = x + 1
            If x = pxWidth Then
                x = 0
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1 '9
            End If
        Loop
End If
End Sub

Public Sub draw1bt(num As Integer, x As Integer, y As Integer)
    If num < 0 Or 1 < num Then
        color1 = RGB(255, 0, 0)
    Else
        c = 255 * (1 - num)
        color1 = RGB(c, c, c)
    End If
    Form1.Picture1.PSet (x, y), color1
End Sub

'##################################################################

Public Sub readImage2bt()
    Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 4 / pxWidth / 320 + 1) * 15 * (pxWidth + 1)
    Dim x As Integer, y As Integer, lx As Integer
    Dim bits0 As Integer
    y = 0
    lx = 0
    If Form1.chkInv.Value = 1 Then
        x = pxWidth - 1
        Do
            bits0 = readbits(2)
            If bits0 = -1 Then Exit Do
            Call draw2bt(bits0, x + lx, y)
            x = x - 1
            If x = -1 Then
                x = pxWidth - 1
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
    Else
        x = 0
        Do
            bits0 = readbits(2)
            If bits0 = -1 Then Exit Do
            Call draw2bt(bits0, x + lx, y)
            x = x + 1
            If x = pxWidth Then
                x = 0
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
    End If
End Sub

Public Sub draw2bt(num As Integer, x As Integer, y As Integer)
    If num < 0 Or 3 < num Then
        color1 = RGB(255, 0, 0)
    Else
        c = 255 - num * 85
        color1 = RGB(c, c, c)
    End If
    Form1.Picture1.PSet (x, y), color1
End Sub

'#####################################################################

Public Sub readImage4bt()
Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 2 / pxWidth / 320 + 1) * 15 * (pxWidth + 1)
    Dim x As Integer, y As Integer, lx As Integer
    Dim bits0 As Integer
    y = 0
    lx = 0
    If Form1.chkInv.Value = 1 Then
        x = pxWidth - 1
        Do
            bits0 = readbits(4)
            If bits0 = -1 Then Exit Do
            Call draw4bt(bits0, x + lx, y)
            x = x - 1
            If x = -1 Then
                x = pxWidth - 1
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
    Else
        x = 0
        Do
            bits0 = readbits(4)
            If bits0 = -1 Then Exit Do
            Call draw4bt(bits0, x + lx, y)
            x = x + 1
            If x = pxWidth Then
                x = 0
                y = y + 1
            End If
            If y = 320 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                y = 0
                lx = lx + pxWidth + 1
            End If
        Loop
End If
End Sub

Public Sub draw4bt(num As Integer, x As Integer, y As Integer)
    If num < 0 Or 15 < num Then
        color1 = RGB(255, 0, 0)
    Else
        c = 255 - num * 17
        color1 = RGB(c, c, c)
    End If
    Form1.Picture1.PSet (x, y), color1
End Sub

'###################################################################

Public Sub readImagePoke()
    Dim y As Integer, lx As Integer
    Dim bits(15) As Integer
    y = 0
    lx = 0
    Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 4 / 8 / 320 + 1) * 15 * (8 + 1)
    trash = readbits(8)
    Do
        For i = 0 To 15
            bits(i) = readbits(1)
            If bits(i) = -1 Then Exit For
        Next i
        For i = 0 To 7
            If bits(i + 8) = -1 Or bits(i) = -1 Then Exit Do
            Call draw2bt(2 * bits(i) + bits(i + 8), (7 - i) + lx, y)
        Next i
        y = y + 1
        If y = 320 Then
            a = DoEvents
            If Not Form1.bStop.Enabled = True Then Exit Do
            y = 0
            lx = lx + 8
        End If
    Loop
End Sub

Public Sub readImagePoke2()
    Dim y As Integer, lx As Integer, ly As Integer
    Dim bits(15) As Integer
    y = 0
    lx = 0
    ly = 0
    Form1.Picture1.Width = 6855
    Form1.Picture1.Height = Int(LOF(1) * 4 / 8 / 456 + 1) * 15 * (8 + 1)
    trash = readbits(8)
    Do
        For i = 0 To 15
            bits(i) = readbits(1)
            If bits(i) = -1 Then Exit For
        Next i
        For i = 0 To 7
            If bits(i + 8) = -1 Or bits(i) = -1 Then Exit Do
            Call draw2bt(2 * bits(i) + bits(i + 8), (7 - i) + lx, y + ly)
        Next i
        y = y + 1
        If y = 8 Then
            y = 0
            lx = lx + 8
            If lx >= 456 Then
                a = DoEvents
                If Not Form1.bStop.Enabled = True Then Exit Do
                lx = 0
                ly = ly + 9
            End If
        End If
    Loop
End Sub

'#############################################################

Public Sub readImage24bt()
    Dim x As Integer, y As Integer, lx As Integer
    Dim bits(2) As Integer
    x = 0
    y = 0
    lx = 0
    Form1.Picture1.Height = 4815
    Form1.Picture1.Width = Int(LOF(1) * 8 / unitBit / pxWidth / 320 + 1) * 15 * (pxWidth + 1)
    If Form1.chkInv.Value = 1 Then
        Do
            For i = 0 To 2
                bits(i) = readbits(8)
                If bits(i) = -1 Then Exit Do
            Next i
            Form1.Picture1.PSet (x + lx, y), RGB(bits(2), bits(1), bits(0))
            x = x - 1
            If x = -1 Then
                x = pxWidth - 1
                y = y + 1
                If y = 320 Then
                    a = DoEvents
                    If Not Form1.bStop.Enabled = True Then Exit Do
                    y = 0
                    lx = lx + pxWidth + 1
                End If
            End If
        Loop
    Else
        Do
            For i = 0 To 2
                bits(i) = readbits(8)
                If bits(i) = -1 Then Exit Do
            Next i
            Form1.Picture1.PSet (x + lx, y), RGB(bits(2), bits(1), bits(0))
            x = x + 1
            If x = pxWidth Then
                x = 0
                y = y + 1
                If y = 320 Then
                    a = DoEvents
                    If Not Form1.bStop.Enabled = True Then Exit Do
                    y = 0
                    lx = lx + pxWidth + 1
                End If
            End If
        Loop
    End If
End Sub

Public Sub draw24bt(r As Integer, g As Integer, b As Integer, x As Integer, y As Integer)
    'If r < 0 Or 255 < r Or g < 0 Or 255 < g Or b < 0 Or 255 < b Then
    '    Picture1.PSet (x, y), RGB(255, 0, 0)
    '    Exit Sub
    'End If
    Form1.Picture1.PSet (x, y), RGB(r, g, b)
End Sub

'################################################################

Public Sub readStringH(filename As String)
    Form1.txtBinDat.Text = ""
    Form1.message openfile(filename)
    a = ""
    Do
        For i = 1 To 8
            b = readBbyH()
            a = a + b + " "
            If b = "EOF" Then Exit Do
        Next i
        Form1.txtBinDat.Text = Form1.txtBinDat.Text + a + Chr(13)
        a = ""
    Loop
    Form1.txtBinDat.Text = Form1.txtBinDat.Text + a + Chr(13)
    Call closefile
End Sub







