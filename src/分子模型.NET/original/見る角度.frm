VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Œ©‚éŠp“x"
   ClientHeight    =   4410
   ClientLeft      =   510
   ClientTop       =   2730
   ClientWidth     =   4410
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown SpinButton1 
      Height          =   735
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      _Version        =   393216
      Max             =   1001
   End
   Begin MSComCtl2.UpDown SpinButton1 
      Height          =   735
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      _Version        =   393216
      Max             =   1001
   End
   Begin MSComCtl2.UpDown SpinButton1 
      Height          =   735
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      _Version        =   393216
      Max             =   1001
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2175
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2175
      Index           =   1
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   2175
      Index           =   0
      Left            =   0
      Shape           =   2  'Oval
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xs(2) As Integer, ys(2) As Integer, xe(2) As Integer, ye(2) As Integer, thrd(2, 2)

Private Sub Form_Load()
For Index = 0 To 2
r = Int(Shape1(Index).Width / 2)
xs(Index) = r + Shape1(Index).Left
ys(Index) = r + Shape1(Index).Top
Next Index
End Sub

Private Sub SpinButton1_Change(Index As Integer)
If SpinButton1(Index).Value = 1001 Then SpinButton1(Index).Value = 1
If SpinButton1(Index).Value = 0 Then SpinButton1(Index).Value = 1000
xe(Index) = Sin(SpinButton1(Index).Value / 500 * 3.1415926535) * xs(Index)
ye(Index) = Cos(SpinButton1(Index).Value / 500 * 3.1415926535) * ys(Index)
Cls
For a = 0 To 2
Line (xs(a), ys(a))-(xe(a), ye(a))
Next a
Dim th(2), cosine(2), sine(2)
For a = 0 To 2
th(a) = SpinButton1(a).Value / 1000 * 3.1415
cosine(a) = Cos(th(a))
sine(a) = Sin(th(a))
Next a
For a = 0 To 2
thrd(a, 1) = thrd(a, 1) * cosine(2) - thrd(a, 2) * sine(2)
thrd(a, 2) = thrd(a, 1) * sine(2) + thrd(a, 2) * cosine(2)
thrd(a, 0) = thrd(a, 0) * cosine(1) - thrd(a, 2) * sine(1)
thrd(a, 2) = thrd(a, 0) * sine(1) + thrd(a, 2) * cosine(1)
thrd(a, 0) = thrd(a, 0) * cosine(0) - thrd(a, 1) * sine(0)
thrd(a, 1) = thrd(a, 0) * sine(0) + thrd(a, 1) * cosine(0)
Next a
Call Form1.ˆÚ“®thrd(thrd(0, 0), thrd(0, 1), thrd(0, 2), thrd(1, 0), thrd(1, 1), thrd(1, 2), thrd(2, 0), thrd(2, 1), thrd(2, 2))
Call Form1.•\Ž¦
End Sub
