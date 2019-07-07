VERSION 5.00
Begin VB.UserControl MathF2 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   Picture         =   "åvéZ2.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   240
End
Attribute VB_Name = "MathF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Type Rational
num As Double
den As Double
sig As Byte
pnt As Integer
End Type

Public Type Complex
Re As Rational
Im As Rational
End Type

'/////Rational/////

Public Function rtarang(number As Rational) As Rational 'ä˘ñÒï™êî
rtarang = number
If rtarang.den = 0 Then
MsgBox "rtarang:ï™ïÍ0Ç≈èúéZ//0Çï‘ÇµÇ‹Ç∑"
rtarang.num = 0
rtarang.den = 1
rtarang.sig = 0
rtarang.pnt = 0
Exit Function
End If
If rtarang.den < 0 Then rtarang.sig = rtarang.sig + 1: rtarang.den = -rtarang.den
If rtarang.num < 0 Then rtarang.sig = rtarang.sig + 1: rtarang.num = -rtarang.num
rtarang.sig = rtarang Mod 2
xx = rtarang.num / 10
While xx = Int(xx)
rtarang.num = xx
rtarang.pnt = rtarang.pnt + 1
xx = xx / 10
Wend
While (rtarang.num * 10) Mod 10
rtarang.num = rtarang.num * 10
rtarang.pnt = rtarang.pnt - 1
Wend
xx = rtarang.den / 10
While xx = Int(xx)
rtarang.num = rtarang.den / 10
rtarang.pnt = rtarang.pnt - 1
xx = xx / 10
Wend
While (rtarang.den * 10) Mod 10
rtarang.pnt = rtarang.pnt + 1
rtarang.den = rtarang.den * 10
Wend
l = LCM(rtarang.num, rtarang.den)
rtarang.num = rtarang.num / l
rtarang.den = rtarang.den / l
End Function

Public Function rtdec(number As Rational) As Double 'rt:dec
On Error GoTo ErrH
rtdec = number.num / number.den * 10 ^ number.pnt
If sig Then rtdec = -rtdec
Exit Function
ErrH:
If Err.number = 6 Then
rtdec = number.num / number.den * (-1) ^ sig
MsgBox "rtdec:µ∞ ﬁ∞Ã€∞//ïÇìÆè¨êîì_e" & number.pnt
End If
End Function

Public Function decrt(number As Double) As Rational  'dec:rt
decrt.num = number
decrt.den = 1
decrt = rtarang(decrt)
End Function

Public Function rtgen(num As Double, den As Double) As Rational 'dec / dec:rt
rtgen.num = num
rtgen.den = den
rtgen = rtarang(rtgen)
End Function

Public Function rtstring(number As Rational) 'rt:"dec / dec"
If number.num = 0 Then rtstring = "0": Exit Function '//
If sig Then rtstring = "-"
a = number.num
b = number.den
If number.pnt < 0 Then b = b & stringprod("0", -pnt)
ElseIf number.pnt > 0 Then
a = a & stringprod("0", pnt)
End If
If b = "1" Then
rtstring = rtstring & a
Else
rtstring = rtstring & a & "/" & b
End If
End Function

Public Function rtsum(num1 As Rational, num2 As Rational) As Rational 'rt + rt
l = LCM(num1.den, num2.den)
rtsum.den = num1.den * num2.den / l
If num1.sig Then num1.num = -num1.num
If num2.sig Then num2.num = -num2.num
num2.num = num2.num * 10 ^ (num1.pnt - num2.pnt)
rtsum.num = num1.num * num2.den / l + num2.num * num1.den / l
rtsum.pnt = num1.pnt
rtsum = rtarang(rtsum)
End Function

Public Function rtdiff(num1 As Rational, num2 As Rational) As Rational 'rt - rt
l = LCM(num1.den, num2.den)
rtdiff.den = num1.den * num2.den / l
If num1.sig Then num1.num = -num1.num
If num2.sig Then num2.num = -num2.num
num2.num = num2.num * 10 ^ (num1.pnt - num2.pnt)
rtdiff.num = num1.num * num2.den / l - num2.num * num1.den / l
rtdiff.pnt = num1.pnt
rtdiff = rtarang(rtdiff)
End Function

Public Function rtprod(num1 As Rational, num2 As Rational) As Rational 'rt * rt
l = LCM(num1.num, num2.den) * LCM(num1.den, num2.num)
rtprod.num = num1.num * num2.num / l
rtprod.den = num1.den * num2.den / l
rtprod.sig = (num1.sig + num2.sig) Mod 2
rtprod.pnt = num1.pnt + num2.pnt
End Function

Public Function rtquot(num1 As Rational, num2 As Rational) As Rational 'rt / rt
l = LCM(num1.num, num2.num) * LCM(num1.den, num2.den)
rtquot.num = num1.num * num2.den / l
rtquot.den = num1.den * num2.num / l
rtquot.sig = (num1.sig + num2.sig) Mod 2
rtquot.pnt = num1.pnt - num2.pnt
End Function

Public Function rtmod(num1 As Rational, num2 As Rational) As Rational 'rt mod rt
rtmod.num = (num1.num / num1.den * 10 ^ num1.pnt) Mod (num2.num / num2.den * 10 ^ num2.pnt)
If num1.sig = 1 Then rtmod.num = (num2.num / num2.den) - rtmod.num
rtmod = rtarang(rtmod)
End Function

Public Function rtbeki(num1 As Rational, num2 As Rational) As Rational 'rt ^ rt
rtbeki.num = num1.num ^ (num2.num / num2.den)
rtbeki.den = num1.den ^ (num2.num / num2.den)
If num1.sig = 1 And num2.num Mod 2 = 1 Then rtbeki.sig = 1
pnt1 = num1.pnt * num2.num / num2.den * 10 ^ num2.pnt
If Int(pnt1) = pnt1 Then rtbeki.pnt = pnt1 Else rtbeki.num = rtbeki.num * 10 ^ pnt1
If num2.sig = 1 Then rtbeki = rtconv(rtbeki)
rtbeki = rtarang(rtbeki)
End Function

Public Function rtconv(number As Rational) As Rational '1 / rt
rtconv.den = number.num
rtconv.num = number.den
rtconv.pnt = -number.pnt
rtconv.sig = number.sig
End Function

Public Function rtabs(number As Rational) As Rational 'abs
rtabs = number
rtabs.sig = 0
End Function

Public Function rtlog(number As Rational) As Rational 'log
Dim a As Rational, num0 As Rational
u = (number.num - number.den) / (number.num + number.den)
uu = u * u
a = decrt(0)
n = 1
num0 = decrt(2.30258509299405 * number.pnt)
Do
num0 = rtsum(num0, a)
a = rtgen(u, n)
n = n + 2
u = u * uu
Loop While rtdec(a)
rtlog = num0
End Function

Public Function rtexp(number As Rational) As Rational 'exp
If number.sig = 1 Then
n = 0: x1 = 1: x2 = 1
Do
rtexp = rtplus(rtexp, rtgen(x1, x2))
n = n + 1
x1 = x1 * number.num
x2 = x2 * number.den * n
While x1 / x2
rtexp = rtconv(rtbeki(rtexp, decrt(10 ^ number.ptn)))
Else
n = 0: x1 = 1: x2 = 1
Do
rtexp = rtplus(rtexp, rtgen(x1, x2))
n = n + 1
x1 = x1 * number.num
x2 = x2 * number.den * n
While x1 / x2
rtexp = rtbeki(rtexp, decrt(10 ^ number.ptn))
End If
End Function

Public Function rtatn(number0 As Rational) As Rational 'atn
Dim number As Rational: number = number0
a = 0
Do While Abs(rtdec(number)) > 0.414213562373095
number = rtquot(rtdiff(decrt(1), number), rtsum(decrt(1), number))
a = a + 1
Loop
rtatn = decrt(0.785398163397448 * a)
a0 = a Mod 2
a = 1: b = decrt(0): nn = number: nn0 = rtprod(number, number)
Do
rtatn = rtsum(rtatn, b)
nn = rtprod(nn, nn0)
a = a + 2
b = rtquot(nn, decrt(a))
If a Mod 4 = 3 Then b.sig = (b.sig + 1 + a0) Mod 2 Else b.sig = (b.sig + a0) Mod 2
Loop
End Function

Public Function rttan(number As Rational) As Rational 'tan
num0 = number.num / number.den * 10 ^ number.pnt
If number.sig Then num0 = -num0
rttan = decrt(Tan(num0))
End Function

Public Function rtsin(number As Rational) As Rational 'sin
num0 = number.num / number.den * 10 ^ number.pnt
If number.sig Then num0 = -num0
rtsin = decrt(Sin(num0))
End Function

Public Function rtcos(number As Rational) As Rational 'cos
num0 = number.num / number.den * 10 ^ number.pnt
If number.sig Then num0 = -num0
rtcos = decrt(Cos(num0))
End Function

'/////Complex/////

Public Function cxgen(Re As Rational, Im As Rational) As Complex 'rt + rt i:cx
cxgen.Re = Re
cxgen.Im = Im
End Function

Public Function rtcx(number As Rational) As Complex 'rt:cx
cxgen.Re = number
cxgen.Im = decrt(0)
End Function

Public Function cxdecgen(Re As Double, Im As Double) As Complex 'dec + dec i:cx
cxdecgen.Re = decrt(Re)
cxdecgen.Im = decrt(Im)
End Function

Public Function deccx(number As Double) As Complex 'dec:cx
cxgen.Re = decrt(number)
cxgen.Im = decrt(0)
End Function

Public Function cxstring(number As Complex) 'cx:"dec / dec + dec / dec i"
cxstring = rtstring(number.Re) & "+" & rtstring(number.Im) & "i"
End Function

Public Function cxconj(number As Complex) As Complex 'cx*
cxconj = number
cxconj.Im.sig = (cxconj.Im.sig + 1) Mod 2
End Function

Public Function cxconv(number As Complex) As Complex '1 / cx
Dim r2 As Rational
r2 = rtsum(rtprod(number.Re, number.Re), rtprod(number.Im, number, Im))
cxconv.Re = rtquot(number.Re, r2)
cxconv.Im = rtquot(rtprod(decrt(-1), number).Im, r2)
End Function

Public Function cxsum(num1 As Complex, num2 As Complex) As Complex 'cx + cx
cxsum.Re = rtsum(num1.Re, num2.Re)
cxsum.Im = rtsum(num1.Im, num2.Im)
End Function

Public Function cxdiff(num1 As Complex, num2 As Complex) As Complex 'cx - cx
cxdiff.Re = rtdiff(num1.Re, num2.Re)
cxdiff.Im = rtdiff(num1.Im, num2.Im)
End Function

Public Function cxprod(num1 As Complex, num2 As Complex) As Complex 'cx * cx
cxprod.Re = rtdiff(rtprod(num1.Re, num2.Re), rtprod(num1.Im, num2.Im))
cxprod.Im = rtsum(rtprod(num1.Re, num2.Im), rtprod(num1.Im, num2.Re))
End Function

Public Function cxquot(num1 As Complex, num2 As Complex) As Complex ' cx / cx
Dim r2 As Rational
r2 = rtsum(rtprod(num2.Re, num2.Re), rtprod(num2.Im, num2.Im))
cxquot.Re = rtquot(rtsum(rtprod(num1.Re, num2.Re), rtprod(num1.Im, num2.Im)), r2)
cxquot.Im = rtquot(rtdiff(rtprod(num1.Im, num2.Re), rtprod(num1.Re, num2.Im)), r2)
End Function

Public Function arg(number As Complex) As Complex 'arg
arg.Re = decrt(Atn(rtdec(number.Im) / rtdec(number.Re)))
arg.Im = 0
End Function

Public Function cxbeki(num As Complex, exp As Complex) As Complex

End Function

Public Function cxsin(number As Complex) As Complex 'sin
cxsin = cxdecgen(Sin(rtdec(number.Re)) * cosh(rtdec(number.Im)), Cos(rtdec(number.Re)) * sinh(rtdec(number.Im)))
End Function

Public Function cxcos(number As Complex) As Complex 'cos
cxcos = cxdecgen(Sin(rtdec(number.Re)) * cosh(rtdec(number.Im)), Cos(rtdec(number.Re)) * sinh(rtdec(number.Im)))
End Function

Public Function cxtan(number As Complex) As Complex 'tan
cxtan = cxdecgen(Sin(2 * rtdec(number.Re)), sinh(2 * rtdec(number.Im)))
cxtan = cxquot(cxtan, deccx(Cos(2 * rtdec(number.Re)) + cosh(2 * rtdec(number.Im))))
End Function

Public Function cxsinh(number As Complex) As Complex 'sinh
cxsinh = cxdecgen(sinh(rtdec(number.Re)) * Cos(rtdec(number.Im)), cosh(rtdec(number.Re)) * Sin(rtdec(number.Im)))
End Function

Public Function cxcosh(number As Complex) As Complex 'cosh
cxcosh = cxdecgen(cosh(rtdec(number.Re)) * Cos(rtdec(number.Im)), sinh(rtdec(number.Re)) * Sin(rtdec(number.Im)))
End Function

Public Function cxtanh(number As Complex) As Complex 'tanh
cxtanh = cxdecgen(sinh(2 * rtdec(number.Re)), Sin(2 * rtdec(number.Im)))
cxtanh = cxquot(cxtanh, deccx(cosh(2 * rtdec(number.Re)) + Cos(2 * rtdec(number.Im))))
End Function

Public Function cxcot(number As Complex) As Complex 'cot
cxcot = cxconv(cxtan(number))
End Function

Public Function cxsec(number As Complex) As Complex 'sec
cxsec = cxconv(cxcos(number))
End Function

Public Function cxcosec(number As Complex) As Complex 'cosec
cxcosec = cxconv(cxsin(number))
End Function

Public Function cxcoth(number As Complex) As Complex 'coth
cxcoth = cxconv(cxtanh(number))
End Function

Public Function cxsech(number As Complex) As Complex 'sech
cxsech = cxconv(cxcosh(number))
End Function

Public Function cxcosech(number As Complex) As Complex 'cosech
cxcosech = cxconv(cxsinh(number))
End Function

'/////Ex/////

Public Function LCM(ByVal num1, ByVal num2) 'ç≈ëÂåˆñÒêî
If num1 = 0 Then LCM = num2: Exit Function
If num2 = 0 Then LCM = num1: Exit Function
If num1 < num2 Then num = num1: num1 = num2: num2 = num
LCM = 1
For a = 2 To num2
While num1 Mod a = 0 And num2 Mod a = 0
LCM = LCM * a
num1 = num1 / a
num2 = num2 / a
Wend
Next a
End Function

Public Function GCM(ByVal num1, ByVal num2) 'ç≈è¨åˆî{êî
GCM = num1 * num / LCM(num1, num2)
End Function

Public Function stringprod(a, n As Integer)
stringprod = ""
If number <= 0 Then Exit Function
For n = 1 To number
stringprod = stringprod & a
Next n
End Function

Public Function sec(number As Double) As Double
sec = 1 / Cos(number)
End Function

Public Function cosec(number As Double) As Double
cosec = 1 / Sin(number)
End Function

Public Function cot(number As Double) As Double
cot = 1 / Tan(number)
End Function

Public Function sinh(number As Double) As Double
sinh = (exp(number) - exp(-number)) / 2
End Function

Public Function cosh(number As Double) As Double
cosh = (exp(number) + exp(-number)) / 2
End Function

Public Function tanh(number As Double) As Double
tanh = (exp(number) - exp(-number)) / (exp(number) + exp(-number))
End Function

Public Function sech(number As Double) As Double
sech = 2 / (exp(number) + exp(-number))
End Function

Public Function cosech(number As Double) As Double
cosech = 2 / (exp(number) - exp(-number))
End Function

Public Function coth()
coth = (exp(number) + exp(-number)) / (exp(number) - exp(-number))
End Function

