VERSION 5.00
Begin VB.UserControl MathF 
   BackColor       =   &H0080FF80&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   Picture         =   "ŒvZctl.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "ŒvZctl.ctx":0342
End
Attribute VB_Name = "MathF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'mPn,mCn,mHn=‡—ñA‘g‡Ad•¡‰Â”\‡—ñ
'fctl=ŠKæ
'loga,log10=‘Î”Aí—p‘Î”
'sec,cosec,cot,arc_sin,arc_cos,arc_cot,arc_sec,arc_cosec=OŠpŠÖ”
'sin_h,cos_h,tan_h,cot_h,sec_h,cosec_h=‘o‹ÈüŠÖ”
'arc_sin_h , arc_cos_h, arc_tan_h, arc_cot_h, arc_sec_h, arc_cosec_h=‹t‘o‹ÈüŠÖ”
'arg=•¡‘f”•ÎŠp
'GCM,LCM=Å‘åŒö–ñ”AÅ¬Œö”{”
'çxdx=x—İæ‚Ìƒ¿‚©‚çƒÀ–˜‚ÌÏ•ª
'ƒÓ=EulerŠÖ”
'Tn,Bn=Bernoulli”
'ƒ¡,ƒÁa,ƒ¡a,Pa,Qa=ƒ¡ŠÖ”
'erf,erfc,nrml_d,nrml_u,ƒÔ2_d,ƒÔ2_u,t_d,t_u,F_d,F_u=GaussŒë·ŠÖ”A[•W€³‹K•ª•z/ƒÔ“ñæ•ª•z/Student‚Ìt•ª•z/‚e•ª•z]‚Ìã‰ºŠm—¦

Public ƒÎ As Double, e As Double, ƒÁ As Double, ƒÓp As Double
Dim lg2p As Double, gammah As Double, rt2pm1 As Double, rt5m1 As Double, ƒÎ2 As Double
Dim ƒÎm1_4 As Double, ƒÎm1_2 As Double, ƒÎh As Double, rtƒÎm1 As Double, ƒÎm1 As Double

Public Function bnml_d(p As Double, n As Integer, k As Integer) As Double
If k < 0 Then
bnml_d = 0
ElseIf k >= n Then
bnml_d = 1
Else
bnml_d = Ix(1 - p, k + 1, n - k)
End If
End Function

Public Function bnml_u(p As Double, n As Integer, k As Integer) As Double
If k <= 0 Then
bnml_u = 1
ElseIf k > n Then
bnml_u = 0
Else
bnml_u = Ix(p, (k), n - k + 1)
End If
End Function

Public Function F_d(ƒË1 As Double, ƒË2 As Double, F As Double) As Double
If F <= 0 Then
F_d = 0
Else
F_d = Ix(ƒË1 / (ƒË1 + ƒË2 / F), ƒË1 / 2, ƒË2 / 2)
End If
End Function

Public Function F_u(ƒË1 As Double, ƒË2 As Double, F As Double) As Double
If F <= 0 Then
F_u = 1
Else
F_u = Ix(ƒË2 / (ƒË2 + ƒË1 / F), ƒË2 / 2, ƒË1 / 2)
End If
End Function

Public Function Ix(x As Double, a As Double, b As Double) As Double
Dim p1 As Double, p2 As Double, q1 As Double, q2 As Double, p As Double, q As Double, r As Integer
Dim pp As Double
p1 = 1
p2 = 1
q1 = 1
q2 = 0
r = 2
Do
C = a + 2 * e
If r Mod 2 = 0 Then
e = Int(r / 2)
d = e * (b - e) / C / (C - 1)
Else
e = Int(r / 2) + a
d = -e * (e + b) / C / (C + 1)
End If
q = x * q2 * d + q1
p1 = p1 / q
p2 = p2 / q
q2 = q1 / q
q1 = 1
pp = p
p = x * p2 * d + p1
p2 = p1
p1 = p
r = r + 1
Loop While pp - p <> 0 And r <= 30000
Ix = x ^ a * (1 - x) ^ b / a / p / ƒ (a, b)
End Function

Public Function Jn(n As Integer, x As Double)
Dim m As Integer, a As Double
Do
Jn = Jn + a
a = (x / 2) ^ (n + 2 * m) / fctl((m)) / fctl(n + m)
If m Mod 2 = 1 Then a = -a
m = m + 1
Loop While a <> 0 And m <= 30000
End Function

Public Function mHn(m As Integer, n As Integer) As Double
mHn = mCn(m + n - 1, n)
End Function

Public Function mCn(m As Integer, n As Integer) As Double
If m >= n Then
mCn = XnFall((m), n) / fctl((n))
Else
mCn = XnFall((n), m) / fctl((m))
MsgBox "mCnŠÖ”‚Ím‚æ‚èn‚Ì•û‚ª‘å‚«‚¢‚½‚ßAm‚Æn‚ğ‹t“]‚µ‚ÄŒvZ‚µ‚Ü‚·B"
End If
End Function

Public Function mPn(m As Integer, n As Integer) As Double
If m >= n Then
mPn = XnFall((m), n)
Else
mPn = XnFall((n), m)
MsgBox "mPnŠÖ”‚Ím‚æ‚èn‚Ì•û‚ª‘å‚«‚¢‚½‚ßAm‚Æn‚ğ‹t“]‚µ‚ÄŒvZ‚µ‚Ü‚·B"
End If
End Function

Public Function fctl(n As Double) As Double
fctl = 1
n = Abs(Int(n))
If n <> 0 Then
For a = 1 To n
fctl = fctl * a
Next a
End If
End Function

Public Function loga(a As Double, C As Double) As Double
loga = Log(C) / Log(a)
End Function

Public Function log10(C As Double) As Double
log10 = Log(C) / Log(10)
End Function

Public Function cosec(r As Double) As Double
cosec = 1 / Sin(r)
End Function

Public Function sec(r As Double) As Double
sec = 1 / Cos(r)
End Function

Public Function cot(r As Double) As Double
cot = 1 / Tan(r)
End Function

Public Function arc_cot(r As Double) As Double
arc_cot = Atn(r) + 2 * Atn(1)
End Function

Public Function arc_sin(r As Double) As Double
arc_sin = Atn(r / Sqr(-r * r + 1))
End Function

Public Function arc_cos(r As Double) As Double
arc_cos = Atn(-r / Sqr(-r * r + 1)) + 2 * Atn(1)
End Function

Public Function arc_sec(r As Double) As Double
arc_sec = Atn(r / Sqr(r * r - 1)) + Sgn((r) - 1) * (2 * Atn(1))
End Function

Public Function arc_cosec(r As Double) As Double
arc_cosec = Atn(r / Sqr(r * r - 1)) + (Sgn(r) - 1) * (2 * Atn(1))
End Function

Public Function sin_h(r As Double) As Double
sin_h = (Exp(r) - Exp(-r)) / 2
End Function

Public Function cos_h(r As Double) As Double
cos_h = (Exp(r) + Exp(-r)) / 2
End Function

Public Function t_d(ƒË As Double, x As Double) As Double
t_d = 1 - Ix(ƒË / (ƒË + t * t), ƒË / 2, 0.5) / 2
End Function

Public Function t_u(ƒË As Double, x As Double) As Double
t_u = Ix(ƒË / (ƒË + t * t), ƒË / 2, 0.5) / 2
End Function

Public Function tan_h(r As Double) As Double
tan_h = (Exp(r) - Exp(-r)) / (Exp(r) + Exp(-r))
End Function

Public Function sec_h(r As Double) As Double
sec_h = 2 / (Exp(r) + Exp(-r))
End Function

Public Function cosec_h(r As Double) As Double
cosec_h = 2 / (Exp(r) - Exp(-r))
End Function

Public Function cot_h(r As Double) As Double
cot_h = (Exp(r) + Exp(-r)) / (Exp(r) - Exp(-r))
End Function

Public Function arc_sin_h(r As Double) As Double
arc_sin_h = Log(r + Sqr(r * r + 1))
End Function

Public Function arc_cos_h(r As Double) As Double
arc_cos_h = Log(r + Sqr(r * r - 1))
End Function

Public Function arc_tan_h(r As Double) As Double
arc_tan_h = Log((1 + r) / (1 - r)) / 2
End Function

Public Function arc_sec_h(r As Double) As Double
arc_sec_h = Log((Sqr(-r * r + 1) + 1) / r)
End Function

Public Function arc_cosec_h(r As Double) As Double
arc_cosec_h = Log((Sgn(r) * Sqr(r * r + 1) + 1) / r)
End Function

Public Function arc_cot_h(r As Double) As Double
arc_cot_h = Log((r + 1) / (r - 1)) / 2
End Function

Public Function arg(real As Double, imag As Double) As Double
arg = Atn(imag / real)
End Function

Public Function çxdx(ƒ¿ As Double, ƒÀ As Double, index As Integer)
index = index + 1
If index > 1 Then index = 1
çxdx = 1 / index * ƒÀ ^ index - 1 / index * ƒ¿ ^ index
End Function

Public Function deg_rad(deg As Double)
deg_rad = deg * ƒÎ / 180
End Function

Public Function rad_deg(rad As Double)
rad_deg = rad * 180 / ƒÎ
End Function

Public Function GCM(a As Double, b As Double) As Double
If a > b Then
a = a + b
b = a - b
a = a - b
End If
For C = 2 To Int(Sqr(a) + 1)
Do While a Mod C = 0 And b Mod C = 0
a = a / C
b = b / C
d = d * C
Loop
Next C
GCM = d
End Function

Public Function LCM(a As Double, b As Double) As Double
LCM = a * b * GCM(a, b)
End Function

Public Function Yn(n As Integer, x As Double) As Double
Dim a As Double, a1 As Double, a2 As Double
Dim s As Integer, b As Double
s = 1
Do
b = Jn(2 * s, x) / s
If s Mod 2 = 0 Then b = -b
a1 = a1 + b
s = s + 1
Loop While b <> 0 And s <= 30000
a1 = a1 * ƒÎm1_4 + Jn(0, x) * (Log(x) + ƒÁ - 0.693147180559945) * ƒÎm1_2
s = 1
b = 0
Do
b = 2 * s + 1
b = Jn((b), x) * b / s / (s + 1)
If s Mod 2 = 0 Then b = -b
a2 = a2 + b
s = s + 1
Loop While b <> 0 And s <= 30000
a2 = (a2 + Jn(1, x) * (Log(x) + ƒÁ - 1.69314718055995) - Jn(0, x) / x) * ƒÎm1_2
If n = 0 Then
Yn = a1
ElseIf n = 1 Then
Yn = a2
ElseIf n > 1 Then
For m = 2 To n
a = 2 * (m - 1) * a2 - a1
a1 = a2
a2 = a
Next m
Yn = a
ElseIf n < 0 Then
For m = -1 To n Step -1
a = 2 * (m - 1) * a1 - a2
a2 = a1
a1 = a
Next m
Yn = a
End If
End Function

Public Function ƒ x(x As Double, a As Double, b As Double) As Double'beta
ƒ x = Ix(x, a, b) * ƒ (a, b)
End Function

Public Function ƒÄ(x As Integer) As Double
If x Mod 2 = 0 Then
ƒÄ = ƒÎ2 ^ x * Abs(Bn(x)) / 2 / fctl((x))
Else
Dim n As Integer, a As Double
n = 1
Do
a = 1 / n ^ x
ƒÄ = ƒÄ + a
n = n + 1
Loop While a <> 0 And n <= 30000
End If
End Function

Public Function ƒÓ(x As Double) As Double
r = x
If x Mod 2 = 0 Then
r = r / 2
Do
x = x / 2
Loop While x Mod 2 = 0
End If
d = 3
Do While x >= d * d
If x Mod d = 0 Then
r = r * (d - 1) / d
Do
x = x / d
Loop While x Mod d = 0
End If
d = d + 2
Loop
ƒÓ = r
End Function

Public Function Tn(n As Integer, x As Double) As Double
Dim ee1 As Double, d As Integer, e1 As Double, e2 As Double
If n <= 0 Then '1
Tn = x
Else '1
 b = "0,1,"
 For a = 0 To n
 For C = 1 To Len(b)
 F = Mid(b, C, 1)
 If F = "," Then '2
 If d > 0 Then ee1 = ee: ee1 = ee1 * d: b1 = b1 & ee1 & ","
 d = d + 1
 ee = ""
 Else '2
 ee = ee & F
 End If '2
 Next C
 d = 0
 For C = 1 To Len(b1) '2
 F = Mid(b1, C, 1)
 If F = "," Then '3
  If d = 1 Then '4
  b = e1 & "," & ee & ","
  Else '4
  b = b & (e + e2) & ","
  End If '4
  d = d + 1
  e2 = e1
  e1 = ee
  ee = ""
 Else '3
  ee = ee & F
 End If '3
 Next C '2
 d = 0
 b = b & e2 & "," & e1 & ","
 Next a
 For C = 1 To Len(b) '2
 F = Mid(b, C, 1)
 If F = "," Then '3
 ee1 = ee
 Tn = Tn + ee1 * x ^ d
 d = d + 1
 ee = ""
 Else '3
 ee = ee & F
 End If '3
 Next C '2
End If '1
End Function
 
Private Sub UserControl_Initialize()
ƒÎ = 3.14159265358979
e = 2.71828182845
ƒÁ = 0.577215664901533
ƒÓp = 1.61803398874989
lg2p = 0.918938533204673
'gammah = ƒ¡(0.5)
rt2pm1 = 0.398942280401433
rt5m1 = 0.447213595499958
ƒÎ2 = 6.28318530717959
ƒÎm1 = 0.318309886183791
ƒÎm1_4 = 0.636619772367581
ƒÎm1_2 = 1.27323954473516
ƒÎh = 1.5707963267949
rtƒÎm1 = 0.564189583547756
End Sub

Public Function Bn(n As Integer) As Double
If n <= 0 Then
Bn = 1
ElseIf n = 1 Then
Bn = -0.5
ElseIf n Mod 2 = 1 Then
Bn = 0
Else
Bn = n * Tn(n - 1, 0)
m = 4 ^ (n / 2)
Bn = Bn / m / (m - 1)
If n Mod 4 = 0 Then Bn = Bn * -1
End If
End Function

Public Function ƒ¡(x As Double) As Double
Dim n As Integer, a As Double, b As Double
n = 2
Do
a = Bn(n) / n / (n - 1) / x ^ (n - 1)
b = b + a
n = n + 2
Loop While n <= 30000 And a <> 0
b = b + lg2p - x + (x - 0.5) * Log(x)
ƒ¡ = Exp(b)
End Function

Public Function XnRise(x As Double, n As Integer) As Double
XnRise = x
If n > 1 Then
For a = x + 1 To x + n - 1 Step 1
XnRise = XnRise * a
Next a
ElseIf n = 0 Then
XnRise = 1
ElseIf n < 0 Then
XnRise = XnFall(x, -n)
End If
End Function

Public Function XnFall(x As Double, n As Integer) As Double
XnFall = x
If n > 1 Then
For a = x - 1 To x - n + 1 Step -1
XnFall = XnFall * a
Next a
ElseIf n = 0 Then
XnFall = 1
ElseIf n < 0 Then
XnFall = XnRise(x, -n)
End If
End Function

Public Function ƒÁa(a As Double, x As Double) As Double
Dim n As Integer
If a >= x + 1 Then
ƒÁa = ƒ¡(x) - ƒ¡a(a, x)
ElseIf a = 0 Then
ƒÁa = 0
Else
Do
b = a ^ n / XnRise(x, n + 1)
n = n + 1
ƒÁa = ƒÁa + b
Loop While b <> 1 And n >= 30000
ƒÁa = ƒÁa * Exp(-a) * a ^ x
End If
End Function

Public Function Ln(n As Integer, p As Double, q As Double) As Double
If n <= 0 Then
Ln = 1
ElseIf n = 1 Then
Ln = p - q + 1
Else
e2 = 1
e1 = p - q + 1
For a = 2 To n
Ln = ((a + p - 1) * (e1 - e2) + (a - q) * e1) / a
e2 = e1
e1 = Ln
Next a
End If
End Function

Public Function ƒ¡a(a As Double, x As Double) As Double
Dim n As Integer
If a < x + 1 Then
ƒ¡a = ƒ¡(x) - ƒÁa(a, x)
Else
Do
b = XnRise(1 - x, n) / fctl(n + 1) / Ln(n, -x, -a) / Ln(n - 1, -x, -a)
n = n + 1
ƒ¡a = ƒ¡a + b
Loop While b <> 0 And n <= 30000
ƒ¡a = ƒ¡a * Exp(-a) * a ^ x
End If
End Function

Public Function Pa(a As Double, x As Double) As Double
Pa = ƒÁa(a, x) / ƒ¡(x)
End Function

Public Function Qa(a As Double, x As Double) As Double
Qa = ƒ¡a(a, x) / ƒ¡(x)
End Function

Public Function erf(x As Double) As Double
If x >= 0 Then
erf = ƒÁa(0.5, x * x) / gammah
Else
erf = -ƒÁa(0.5, x * x) / gammah
End If
End Function

Public Function erf_c(x As Double) As Double
If x >= 0 Then
erf_c = ƒ¡a(0.5, x * x) / gammah
Else
erf_c = 1 + ƒÁa(0.5, x * x) / gammah
End If
End Function

Public Function nrml_d(x As Double) As Double
If x >= 0 Then
nrml_d = 0.5 + ƒÁa(0.5, x * x) / gammah / 2
Else
nrml_d = nrml_u(-x)
End If
End Function

Public Function nrml_u(x As Double) As Double
If x >= 0 Then
nrml_u = 0.5 - ƒÁa(0.5, x * x) / gammah / 2
Else
nrml_u = nrml_d(-x)
End If
End Function

Public Function ƒÔ2_d(ƒË As Double, ƒÔ2 As Double) As Double
ƒÔ2_d = ƒÁa(ƒË / 2, ƒÔ2 / 2) / ƒ¡(ƒË / 2)
End Function

Public Function ƒÔ2_u(ƒË As Double, ƒÔ2 As Double) As Double
ƒÔ2_u = ƒ¡a(ƒË / 2, ƒÔ2 / 2) / ƒ¡(ƒË / 2)
End Function

Public Function Ea(a As Double, x As Double) As Double
Ea = ƒ¡a(1 - a, x) * x ^ (a - 1)
End Function

Public Function bnml(p As Double, n As Integer, k As Integer) As Double
bnml = mCn((n), (k)) * p ^ k * (1 - p) ^ (n - k)
End Function

Public Function nrml(m As Double, s As Double, x As Double) As Double
a = m - x
nmrl = rt2pm1 / s / Exp(-a * a / 2 / s / s)
End Function

Public Function erf_l(s As Double, x As Double) As Double
erf_l = rt2pm1 / s / Exp(-x * x / 2 / s / s)
End Function

Public Function ƒÕ(x As Double) As Double
Dim n As Integer, a As Double
n = 2
Do
ƒÕ = ƒÕ - a
a = Bn(n) / n / x ^ n
n = n + 2
Loop While a <> 0 And n <= 30000
ƒÕ = ƒÕ + Log(x) - 0.5 / x
End Function

Public Function nrml_l(x As Double) As Double
nrml_l = rt2pm1 / Exp(-x * x / 2)
End Function

Public Function ƒÕn(n As Integer, x As Double) As Double
Dim m As Integer, a As Double, xn As Double
m = 2
xx = x * x
xn = xx
Do
ƒÕn = ƒÕn + a
a = Bn(m) * XnRise(m + 1, n - 1) / xn
xn = xn * xx
m = m + 2
Loop While m <> 0 And m <= 30000
ƒÕn = (ƒÕn + fctl(n - 1) + fctl((n)) / 2) / x ^ n
If n Mod 2 = 0 Then ƒÕn = -ƒÕn
End Function

Public Function ƒ (a As Double, b As Double)
Dim n As Integer, d As Double, p As Double
n = 1
C = a + b
Do
d = d + p
p = (1 / a ^ n + 1 / b ^ n - 1 / C ^ n) * Bn(n + 1) / n / (n + 1)
n = n + 2
Loop While p <> 0 And n <= 30000
d = d + lg2p + Log(a) * (a - 1) + Log(b) * (b - 1) - Log(C) * (C - 1)
ƒ  = Exp(d)
End Function

Public Function Fn(n As Integer) As Double
Fn = Int(rt5m1 * ƒÓp ^ n + 0.5)
End Function

Public Function Geo_den(p As Double, n As Integer) As Double
Geo_den = p * (1 - p) ^ (n - 1)
End Function

Public Function ƒ¡_den(a As Double, x As Double) As Double
If a < 0 Then a = -a
ƒ¡_den = x ^ (a - 1) / ƒ¡(x) / Exp(x)
End Function

Public Function ƒÔ2_den(ƒË As Integer, ƒÔ As Double) As Double
ƒË = ƒË / 2
If ƒÔ < 0 Then ƒÔ = -ƒÔ
ƒÔ2_den = ƒÔ ^ (ƒË - 1) * Exp(-ƒÔ / 2) / 2 ^ ƒË / ƒ¡((ƒË))
End Function

Public Function Tri_den(x As Double) As Double
If x < -1 Then
x = -1
ElseIf x > 1 Then
x = 1
End If
Tri_den = 1 - Abs(x)
End Function

Public Function Xp_den(x As Double) As Double
If x < 0 Then x = -x
Xp_den = Exp(-x)
End Function

Public Function F_den(ƒË1 As Double, ƒË2 As Double, x As Double) As Double
Dim a As Double, b As Double, C As Double
a = ƒË1 / 2
b = ƒË2 / 2
C = x * ƒË1 / ƒË2
F_den = C ^ a / (1 + C) ^ (a + b) / x / ƒ (a, b)
End Function

Public Function Si(x As Double) As Double
Dim a As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, x2 As Double, x4 As Double
x2 = x * x
x4 = x2 * x2
a = 1
b1 = x
b2 = 1
b3 = 18
b4 = x2
Do
C = b1 * (b3 - b4) / b2 / a / b3
Si = Si + C
a = a + 4
b1 = b1 * x4
b2 = b2 * XnFall(a, 4)
b3 = a + 2
b3 = b3 * b3 * (a + 1)
b4 = x2 * a
Loop While C <> 0 And a <= 30000
End Function

Public Function si2(x As Double) As Double
si2 = Si(x) - ƒÎh
End Function

Public Function Ci(x As Double) As Double
Dim a As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, x2 As Double, x4 As Double
x2 = x * x
x4 = x2 * x2
a = 4
b1 = x4
b2 = 24
b3 = 180
b4 = 4 * x2
Do
C = b1 * (b3 - b4) / b2 / a / b3
Ci = Ci + C
a = a + 4
b1 = b1 * x4
b2 = b2 * XnFall(a, 4)
b3 = a + 2
b3 = b3 * b3 * (a + 1)
b4 = x2 * a
Loop While C <> 0 And a <= 30000
Ci = Ci + ƒÁ + Log(x) - x ^ 4 / 4
End Function

Public Function Lgs_den(x As Double) As Double
a = Exp(-x)
b = a + 1
Lgs_den = a / b / b
End Function

Public Function Lgs_l(x As Double) As Double
Lgs_l = 1 / (1 + Exp(-x))
End Function

Public Function Ix_den(x As Double, a As Double, b As Double)
If a < 0 Then
a = -a
ElseIf a = 0 Then
a = 1
End If
If b < 0 Then
b = -b
ElseIf b = 0 Then
b = 1
End If
If x < 0 Then
x = 0
ElseIf x > 1 Then
x = 1
End If
Ix_den = x ^ (a - 1) * (1 - x) ^ (b - 1) / Ix(x, a, b)
End Function

Public Function Wei_den(ƒ¿ As Double, x As Double) As Double
Wei_den = 1 - Exp(-x ^ ƒ¿)
End Function

Public Function t_den(ƒË As Integer, x As Double) As Double
t_den = ƒ¡((ƒË + 1) / 2) * (1 + x * x / ƒË) ^ ((ƒË + 1) / -2) * rtƒÎm1 / ƒ¡(ƒË / 2) / Sqr(ƒË)
End Function

Public Function Cach_den(x As Double) As Double
Cach_den = ƒÎm1 * (1 + x * x)
End Function

Public Function mE1n(m As Integer, n As Integer)
If n = 0 Then
mE1n = 0
ElseIf n < 0 Or n >= m Then
mE1n = 1
Else
mE1n = (n + 1) * mE1n(m - 1, n) + (m - n) * mE1n(m - 1, n - 1)
End If
End Function

Public Function mE2n(m As Integer, n As Integer) As Double
n = n - 1
If n = 0 Then
mE2n = 0
ElseIf n < 0 Or n >= m Then
mE2n = 1
Else
mE2n = (n + 1) * mE2n(m - 1, n) + (m - n) * mE2n(m - 1, n - 1)
End If
End Function

Public Function mS1n(m As Integer, n As Integer) As Integer
If n < 1 Or n > m Then
mS1n = 0
ElseIf n = m Then
mS1n = 1
Else
mS1n = (m - 1) * mS1n(m - 1, n) + mS1n(m - 1, n - 1)
End If
End Function

Public Function mS2n(m As Integer, n As Integer) As Double
If n < 1 Or n > m Then
mS2n = 0
ElseIf n = 1 Or n = m Then
mS2n = 1
Else
mS2n = n * mS2n(m - 1, n) + mS2n(m - 1, n - 1)
End If
End Function

