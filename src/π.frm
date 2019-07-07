VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ÉŒ"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "ÉNÉäÉA"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   2640
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   10
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   8
      Text            =   "0"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ωƒ∞œ∞åvéZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "óêêîåvéZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "è°ñ⁄åvéZÇQ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "è°ñ⁄åvéZÇP"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ëΩäpå`åvéZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000FF00&
      Caption         =   "åvéZíÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "åvéZë¨ìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "å¬"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "è°"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "è°"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "äpå`"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Å@Å@åvéZó "
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label14.Visible = True
Label12.Caption = 0
b = Text1.Text
c = 1
Timer1.Interval = 1000
For a = 1 To b
c = 2 - Sqr(4 - c)
Next a
Timer1.Interval = 0
Label1.Caption = Sqr(c) * 3 * 2 ^ b
Label14.Visible = False
End Sub

Private Sub Command2_Click()
Label14.Visible = True
Label12.Caption = 0
Dim e As Single
a = Text2.Text
b = a ^ 2
c = 0
d = a
e = b
Timer1.Interval = 1000
For f = 0 To a
Do Until e <= b
e = e - d * 2 + 1
d = d - 1
Loop
c = c + d
e = e + f * 2 + 1
Next f
Timer1.Interval = 0
Label1.Caption = (c * 4 + 1) / b
Label14.Visible = False
End Sub

Private Sub Command3_Click()
Label14.Visible = True
Label12.Caption = 0
a = Text3.Text
b = a ^ 2
c = 0
Timer1.Interval = 1000
For d = 1 To a
For e = 0 To a
If d ^ 2 + e ^ 2 < -b Then c = c + 1
Next e
Next d
Timer1.Interval = 0
Label1.Caption = c * 4 + 1
Label14.Visible = False
End Sub

Private Sub Command4_Click()
Label14.Visible = True
Label12.Caption = 0
Dim e As Currency
a = Text4.Text
e = 0
Randomize
Timer1.Interval = 1000
For x = 1 To a
For b = 1 To 10000
c = Rnd
d = Rnd
If c ^ 2 + d ^ 2 < 1 Then e = e + 1
Next b
Next x
Timer1.Interval = 0
Label1.Caption = 4 * e / a / 10000
Label14.Visible = False
End Sub

Private Sub Label14_Click()
Label14.Visible = True
Label12.Caption = 0
Timer1.Interval = 1000
Timer1.Interval = 0
Label14.Visible = False
End Sub

Private Sub Text1_Change()
On Error GoTo error
Label4.Caption = 3 * 2 ^ Text1.Text
Exit Sub
error: Label5.Caption = 0
End Sub

Private Sub Text2_Change()
On Error GoTo error
Label5.Caption = Text2.Text ^ 2
Exit Sub
error: Label5.Caption = 0
End Sub

Private Sub Text3_Change()
On Error GoTo error
Label7.Caption = Text3.Text ^ 2
Exit Sub
error: Label5.Caption = 0
End Sub

Private Sub Text4_Change()
On Error GoTo error
Label8.Caption = Text4.Text * 10000
Exit Sub
error: Label5.Caption = 0
End Sub

Private Sub Timer1_Timer()
Dim a As Integer, b As Integer, c As Integer
a = 1: b = Label12.Caption: c = a + b: Label12.Caption = c
End Sub
