VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "é¿å±"
   ClientHeight    =   8655
   ClientLeft      =   2700
   ClientTop       =   1230
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   9810
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "form2.show"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ëSêFï\é¶"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "êFï\é¶ 255;1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   120
      Max             =   8
      Min             =   1
      TabIndex        =   3
      Top             =   7800
      Value           =   1
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   120
      Max             =   256
      Min             =   2
      SmallChange     =   2
      TabIndex        =   2
      Top             =   7680
      Value           =   255
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   7680
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acn As Integer
Private Sub Command1_Click()
Open "c:\windows\√ﬁΩ∏ƒØÃﬂ\è[\Excel2000\Excel2000Åiè[Åj.xls" For Input As 1
Input #1, a
Close #1
MsgBox a
End Sub

Private Sub Command2_Click()
Form1.Cls
acn = -1
stp0 = HScroll1.Value
stp1 = HScroll2.Value
stp = 1 / stp0
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(a, a, a), stp1)
Next a
Form1.Caption = "é¿å±-åvéZíÜÅ†Å†Å†Å†Å†Å†"
For b = 0 To 1 - stp Step stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(a, a * b, 0), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ†Å†Å†Å†Å†Å°"
For b = 1 To stp Step -stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(a * b, a, 0), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ†Å†Å†Å†Å°Å°"
For b = 0 To 1 - stp Step stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(0, a, a * b), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ†Å†Å†Å°Å°Å°"
For b = 1 To stp Step -stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(0, a * b, a), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ†Å†Å°Å°Å°Å°"
For b = 0 To 1 - stp Step stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(a * b, 0, a), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ†Å°Å°Å°Å°Å°"
For b = 1 To stp Step -stp
For a = 0 To 255 Step stp1
Call COLOR(a, RGB(a, 0, a * b), stp1)
Next a
Next b
Form1.Caption = "é¿å±-åvéZíÜÅ°Å°Å°Å°Å°Å°"
MsgBox "äÆóπ", 64
Form1.Caption = "é¿å±"
End Sub

Private Sub Command3_Click()
Form1.Cls
stp0 = 4
For a = 0 To 255 Step stp0
Form1.Caption = "é¿å±-åvéZíÜ" & Int(a / 2.55) & "%äÆóπ"
aa = a / stp0 * 15
For b = 0 To 255 Step stp0
bb = b / stp0 * 15
For c = 0 To 255 Step stp0
Form1.PSet (aa + (c / stp0 * 256 / stp0 + 1) * 15 - Int(c / 32) * 256 / stp0 * 120, bb + Int(c / 32) * 256 / stp0 * 15), RGB(a, b, c)
Next c
Next b
Next a
MsgBox "äÆóπ"
Form1.Caption = "é¿å±"
End Sub

Private Sub Command4_Click()
Form2.Show
End Sub

Private Sub Command5_Click()
Open "c:\é¿å±ÇP" For Output As 1
For a = 0 To 255
b = b & ChrW(&HE000& + a)
Next a
Write #1, b
Close #1
Open "c:\é¿å±ÇP" For Input As 1
Input #1, b
Close #1
For a = 1 To 256
c = c & AscW(Mid(b, a, 1)) + &H2000& & ","
Next a
Kill "c:\é¿å±ÇP"
MsgBox c & Chr(13) & b
End Sub

Private Sub Form_Load()
'MsgBox "hi"
'MsgBox "hi", 16
'MsgBox "hi", 32
'MsgBox "hi", 48
'MsgBox "hi", 64
End Sub

Public Sub COLOR(a, col, stp1)
If a = 0 Then acn = acn + 1
xx = a / stp1 * 15 + Int(acn / 512) * 15 * ((256 - stp1) / stp1 + 2)
yy = acn * 15 - Int(acn / 512) * 15 * 512
Form1.PSet (xx, yy), col
End Sub

Private Sub HScroll1_Change()
Command2.Caption = "êFï\é¶ " & HScroll1.Value & ";" & HScroll2.Value
End Sub

Private Sub HScroll2_Change()
Command2.Caption = "êFï\é¶ " & HScroll1.Value & ";" & HScroll2.Value
End Sub
