VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "ÇœÇÁÇœÇÁ"
   ClientHeight    =   3825
   ClientLeft      =   5445
   ClientTop       =   4590
   ClientWidth     =   4680
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll5 
      Height          =   375
      Left            =   1320
      Max             =   99
      Min             =   39
      TabIndex        =   15
      Top             =   3360
      Value           =   99
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "èâä˙âª"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "èIóπ"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ì«çû"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ï€ë∂"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      Left            =   3000
      Max             =   1000
      Min             =   1
      TabIndex        =   10
      Top             =   2760
      Value           =   100
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "í‚é~"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "çƒê∂"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   840
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3480
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3480
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "çƒê∂ë¨ìx100"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   6.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2880
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1560
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin MSForms.SpinButton SpinButton2 
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   480
      Width           =   255
      Size            =   "450;450"
      Max             =   99
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   255
      Size            =   "450;450"
      Max             =   99
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(99, 99, 99), a1 As Integer, a2(99, 99), a3 As Integer, tate, yoko

Private Sub Command1_Click()
Timer1.Interval = HScroll4.Value
Form1.Caption = "çƒê∂íÜ"
End Sub

Private Sub Command2_Click()
Timer1.Interval = 0
Form1.Caption = "í‚é~íÜ"
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowSave
aa = CommonDialog1.filename
If aa = "" Then Exit Sub
Open aa For Output As 1
For aa0 = 0 To 99
Form1.Caption = "ï€ë∂íÜ-" & aa0 & "%äÆóπ"
For aa1 = 0 To 99
bbb = a(aa0, aa1, 0)
bb0 = Int(bbb / 256 ^ 2)
bbb = bbb - bb0 * 256 ^ 2
bb1 = Int(bbb / 256)
bbb = bbb - bb1 * 256
bb2 = bbb
bb = ChrW(&HE000& + bb0) & ChrW(&HE000& + bb1) & ChrW(&HE000& + bb2)
For aa2 = 1 To 99
bbb = a(aa0, aa1, aa2)
bb0 = Int(bbb / 256 ^ 2)
bbb = bbb - bb0 * 256 ^ 2
bb1 = Int(bbb / 256)
bbb = bbb - bb1 * 256
bb2 = bbb
bb = bb & ChrW(&HE000& + bb0) & ChrW(&HE000& + bb1) & ChrW(&HE000& + bb2)
Next aa2
Print #1, bb
Next aa1
Next aa0
Close #1
Form1.Caption = "í‚é~íÜ"
End Sub

Private Sub Command4_Click()
CommonDialog1.ShowOpen
aa = CommonDialog1.filename
If aa = "" Then Exit Sub
Open aa For Input As 1
For aa0 = 0 To 99
For aa1 = 0 To 39
For aa2 = 0 To 39
Input #1, bb
a(aa0, aa1, aa2) = bb
Next aa2
Next aa1
Next aa0
Close #1
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Form1.Cls
tate = HScroll5.Value
yoko = HScroll5.Value
Call shokika
End Sub

Private Sub Form_Load()
tate = 99
yoko = 99
Call shokika
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If X <= tate * 15 And X >= 0 And Y <= yoko * 15 And Y >= 0 Then
a3 = 1
ElseIf X <= tate * 15 And X >= 0 And Y >= (yoko + 2) * 15 And Y <= (2 * yoko + 2) * 15 Then
a3 = 2
End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call pointè[(X, Y)
Call pointè[(X + 15, Y)
Call pointè[(X, Y + 15)
Call pointè[(X + 15, Y + 15)
End If
If X <= 15 * tate And Y <= 15 * yoko And Text1.Text = "" Then Text1.Text = 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If X <= tate * 15 And X >= 0 And Y >= (yoko + 2) * 15 And Y <= (2 * yoko + 2) * 15 And a3 = 1 Then
For b = 0 To tate
For c = 0 To yoko
a2(b, c) = a(Text1.Text, b, c)
Next c
Next b
Call azuke
ElseIf X <= tate * 15 And X >= 0 And Y <= yoko * 15 And Y >= 0 And a3 = 2 Then
For b = 0 To tate
For c = 0 To yoko
a(Text1.Text, b, c) = a2(b, c)
Next c
Next b
Call Text1_Change
End If
a3 = 0
End If
End Sub

Private Sub HScroll1_Change()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Call HScroll1_Change
End Sub

Private Sub HScroll3_Change()
Call HScroll1_Change
End Sub

Public Sub pointè[(X As Single, Y As Single)
If Text1.Text = "" Then Text1.Text = 0
xx = X / 15
If xx > tate Or xx < 0 Then GoTo err
yy = Y / 15
If yy > yoko Or yy < 0 Then GoTo err
a(Text1.Text, xx, yy) = Picture1.BackColor
Form1.PSet (X, Y), Picture1.BackColor
err:
End Sub

Private Sub HScroll4_Change()
Label2.Caption = "çƒê∂ë¨ìx" & HScroll4.Value
End Sub

Private Sub HScroll5_Change()
Label1.Caption = HScroll5.Value
End Sub

Private Sub SpinButton1_Change()
Text1.Text = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
Text2.Text = SpinButton2.Value
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then GoTo err1
Dim d As Integer
On Error GoTo err2
d = Text1.Text
If d > 99 Or d < 0 Then GoTo err2
Form1.Caption = "ï\é¶íÜ"
For b = 0 To tate
For c = 0 To yoko
Form1.PSet (b * 15, c * 15), a(d, b, c)
Next c
Next b
Form1.Caption = "í‚é~íÜ"
SpinButton1.Value = Text1.Text
Exit Sub
err1:
Exit Sub
err2:
MsgBox "0Ç©ÇÁ99ñòÇÃêîéöÇêÆêîÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅI"
Text1.Text = 0
Form1.Caption = "í‚é~íÜ"
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then Text1.Text = 0
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then GoTo err1
On Error GoTo err2:
SpinButton2.Value = Text2.Text
err1:
Exit Sub
err2:
MsgBox "0Ç©ÇÁ99ñòÇÃêîéöÇêÆêîÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅI"
Text2.Text = 0
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then Text2.Text = 0
End Sub

Private Sub Timer1_Timer()
For b = 0 To tate
For c = 0 To yoko
Form1.PSet (b * 15, c * 15), a(a1, b, c)
Next c
Next b
a1 = a1 + 1
If a1 > Text2.Text Then a1 = 0
End Sub

Public Sub azuke()
For b = 0 To tate
For c = 0 To yoko
Form1.PSet (b * 15, (c + yoko + 2) * 15), a2(b, c)
Next c
Next b
End Sub

Public Sub shokika()
s = RGB(0, 0, 0)
For b = 0 To 99
For c = 0 To tate
For d = 0 To yoko
a(b, c, d) = s
Next d
Next c
Next b
Call Text1_Change
Call azuke
End Sub
