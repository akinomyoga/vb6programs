VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "設定..."
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   958
      Left            =   4320
      Top             =   3720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "HGP行書体"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "HGP行書体"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hms(2), ymd(2), r1, r2, ox(2, 1), oy(2, 1), col(6, 2)
Public ohr, omr, osr

Private Sub Command1_Click()
Form2.Show
For a = 0 To 6
For b = 0 To 2
Call Form2.colpass((a), (b), (col(a, b)))
Next b
Next a
End Sub

Private Sub Form_Load()
r2 = Picture1.Height / 2
r1 = Picture1.Width / 2
For a = 3 To 5
For b = 0 To 2
col(a, b) = 255
Next b
Next a
For b = 0 To 2
col(6, b) = 224
Next b
End Sub

Private Sub Timer1_Timer()
anow = Now
Form1.Caption = anow
adate = Date
b = 0
c = ""
For a = 1 To 10
If Mid(adate, a, 1) = "/" Then
ymd(b) = c
b = b + 1
c = ""
If b = 3 Then Exit For
Else
c = c & Mid(adate, a, 1)
End If
Next a
ymd(2) = c
atime = Time
b = 0
c = ""
For a = 1 To 8
If Mid(atime, a, 1) = ":" Then
hms(b) = c
b = b + 1
c = ""
If b = 3 Then Exit For
Else
c = c & Mid(atime, a, 1)
End If
Next a
hms(2) = c
If hms(0) >= 12 Then
ampm = 1
hms(0) = hms(0) - 12
End If
hr = hms(0) / 6 * 3.1416
If hr <> ohr Then
ohr = hr
hx = Sin(hr)
hy = -Cos(hr)
a = r1
b = r2
c = r1 + hx * r1 / 3
d = r2 + hy * r2 / 3
Picture1.DrawWidth = 3
Picture1.Line (a, b)-(c, d), RGB(col(0, 0), col(0, 1), col(0, 2))
Picture1.Line (ox(0, 0), oy(0, 0))-(ox(0, 1), oy(0, 1)), RGB(col(5, 0), col(5, 1), col(5, 2))
ox(0, 0) = a
oy(0, 0) = b
ox(0, 1) = c
oy(0, 1) = d
End If
mr = hms(1) / 30 * 3.1416
If mr <> omr Then
omr = mr
mx = Sin(mr)
my = -Cos(mr)
a = r1 + mx * r1 / 3
b = r2 + my * r2 / 3
c = r1 + mx * r1 / 3 * 2
d = r2 + my * r2 / 3 * 2
Picture1.DrawWidth = 2
Picture1.Line (a, b)-(c, d), RGB(col(1, 0), col(1, 1), col(1, 2))
Picture1.Line (ox(1, 0), oy(1, 0))-(ox(1, 1), oy(1, 1)), RGB(col(4, 0), col(4, 1), col(4, 2))
ox(1, 0) = a
oy(1, 0) = b
ox(1, 1) = c
oy(1, 1) = d
End If
sr = hms(2) / 30 * 3.1416
If sr <> osr Then
osr = sr
sx = Sin(sr)
sy = -Cos(sr)
a = r1 + r1 * sx / 3 * 2
b = r2 + r2 * sy / 3 * 2
c = r1 + r1 * sx
d = r2 + r2 * sy
Picture1.DrawWidth = 1
Picture1.Line (a, b)-(c, d), RGB(col(2, 0), col(2, 1), col(2, 2))
Picture1.Line (ox(2, 0), oy(2, 0))-(ox(2, 1), oy(2, 1)), RGB(col(3, 0), col(3, 1), col(3, 2))
ox(2, 0) = a
oy(2, 0) = b
ox(2, 1) = c
oy(2, 1) = d
End If
Label1.Caption = "西暦" & ymd(0) & "年" & ymd(1) & "月" & ymd(2) & "日"
a = "午前"
If ampm = 1 Then a = "午後"
Label2.Caption = a & hms(0) & "時" & hms(1) & "分" & hms(2) & "秒"
End Sub

Public Sub colpass(a As Integer, b As Integer, c As Integer)
col(a, b) = c
End Sub
