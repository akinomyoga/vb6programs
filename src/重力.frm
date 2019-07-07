VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "èdóÕ"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "ãOìπ"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "ãOê’ÇécÇ∑"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "è¡Ç∑"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ãÖï\é¶"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   15.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Å°"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "é¿çsÇR"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "é¿çsÇQ"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "é¿çs"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin MSComCtl2.UpDown SpinButton2 
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   8400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Value           =   1
      Max             =   50
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown SpinButton1 
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   8040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Value           =   3
      Max             =   20
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "ï\é¶âÒêî1/1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   3240
      TabIndex        =   10
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "ë¨ìx 3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   8040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx(1, 80) As Double, yy(1, 80) As Double, zz(1, 80) As Double
Dim ox As Integer, oy As Integer

Public Sub drawpoint(x As Double, y As Double, z As Double)
    X1 = (x - y) * Sqr(3) / 2 + ox
    Y1 = z - (x + y) / 2 + oy
    PSet (X1, Y1), Form1.ForeColor
End Sub

Private Sub Command1_Click()
    For a1 = 0 To 8
        For a2 = 0 To 8
            xx(0, a1 * 9 + a2) = a2 * 100
            yy(0, a1 * 9 + a2) = 2000
            zz(0, a1 * 9 + a2) = a1 * 100
            xx(1, a1 * 9 + a2) = 0
            yy(1, a1 * 9 + a2) = -SpinButton1.Value
            zz(1, a1 * 9 + a2) = 0
        Next a2
    Next a1
    Timer1.Interval = 1
End Sub

Private Sub Command2_Click()
    Cls
End Sub

Private Sub Command3_Click()
    a1 = 300 'ãÖÇÃëÂÇ´Ç≥
    a2 = 15 'çèÇ›íl
    For a = 0 To 90 Step a2
        z = a1 * Sin(a / 180 * 3.14159265)
        c = a1 * Cos(a / 180 * 3.14159265)
        For b = 0 To 90 Step a2
            x = c * Cos(b / 180 * 3.14159265)
            y = c * Sin(b / 180 * 3.14159265)
            Call drawpoint(x * 1, y * 1, z * 1)
            Call drawpoint(x * 1, y * 1, z * -1)
            Call drawpoint(x * 1, y * -1, z * 1)
            Call drawpoint(x * 1, y * -1, z * -1)
            Call drawpoint(x * -1, y * 1, z * 1)
            Call drawpoint(x * -1, y * 1, z * -1)
            Call drawpoint(x * -1, y * -1, z * 1)
            Call drawpoint(x * -1, y * -1, z * -1)
        Next b
    Next a
End Sub

Private Sub Command4_Click()
    Timer1.Interval = 0
End Sub

Private Sub Command5_Click()
    For a1 = 0 To 80
        xx(0, a1) = -1000
        yy(0, a1) = 1000
        zz(0, a1) = a1 * 30 + 500
        xx(1, a1) = SpinButton1.Value
        yy(1, a1) = -SpinButton1.Value
        zz(1, a1) = 0
    Next a1
    Timer1.Interval = 1
End Sub

Private Sub Command6_Click()
    For a1 = 0 To 80
        xx(0, a1) = 0
        yy(0, a1) = 0
        zz(0, a1) = a1 * 30 + 500
        xx(1, a1) = SpinButton1.Value
        yy(1, a1) = -SpinButton1.Value
        zz(1, a1) = 0
    Next a1
    Timer1.Interval = 1
End Sub

Private Sub Command7_Click()
    xx(0, a1) = 200
    yy(0, a1) = 200
    zz(0, a1) = 2800
    xx(1, a1) = 2
    yy(1, a1) = -2
    zz(1, a1) = 0
    Timer1.Interval = 1
End Sub

Private Sub Command8_Click()
    Timer1.Interval = 1
End Sub

Private Sub Form_Load()
    Form2.Show
    ox = 7200
    oy = 4800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub SpinButton1_Change()
    Label1.Caption = "ë¨ìx " & SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    Label2.Caption = "ï\é¶âÒêî1/" & SpinButton2.Value
End Sub

Private Sub Timer1_Timer()
    b2 = Form2.ScaleHeight
    For b1 = 1 To SpinButton2.Value
        Form2.Cls
        For a1 = 0 To 80
            a2 = Sqr(xx(0, a1) ^ 2 + yy(0, a1) ^ 2 + zz(0, a1) ^ 2) 'distance
            a3 = Sqr(xx(1, a1) ^ 2 + yy(1, a1) ^ 2 + zz(1, a1) ^ 2) 'ë¨Ç≥
            Form2.Line (a1 * 15, b2)-(a1 * 15, b2 - a3 * 75)
            If a2 <= 300 Then
                ForeColor = RGB(255, 0, 0)
                Call drawcircle(xx(0, a1), yy(0, a1), zz(0, a1))
                ForeColor = RGB(255, 255, 255)
                xx(0, a1) = 0
                yy(0, a1) = 0
                zz(0, a1) = 0
                xx(1, a1) = 0
                yy(1, a1) = 0
                zz(1, a1) = 0
            End If
            On Error GoTo Err
            If a2 <> 0 Then
                a3 = 1 / a2 ^ 2.1 / a2 * 100000 'ãóó£Ç…ëŒÇ∑ÇÈë¨ìxïœâªó ÇÃäÑçá
                xx(1, a1) = xx(1, a1) - xx(0, a1) * a3
                yy(1, a1) = yy(1, a1) - yy(0, a1) * a3
                zz(1, a1) = zz(1, a1) - zz(0, a1) * a3
            End If
Err:
            xx(0, a1) = xx(0, a1) + xx(1, a1)
            yy(0, a1) = yy(0, a1) + yy(1, a1)
            zz(0, a1) = zz(0, a1) + zz(1, a1)
        Next a1
    Next b1
    If Check1.Value = 0 Then Cls
    For a1 = 0 To 80
        Call drawpoint(xx(0, a1), yy(0, a1), zz(0, a1))
    Next a1
    Call Command3_Click
End Sub

Public Sub drawcircle(x As Double, y As Double, z As Double)
    X1 = (x - y) * Sqr(3) / 2 + ox
    Y1 = z - (x + y) / 2 + oy
    Circle (X1, Y1), 150, Form1.ForeColor
End Sub
