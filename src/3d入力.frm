VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "ì¸óÕ"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.OptionButton Option1 
      Caption         =   "íºï˚ëÃ"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   15
      Top             =   4800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "íºê¸"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   14
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OKÅiï\é¶âÊñ Ç…èÓïÒÇëóÇÈÅj"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   7560
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Index           =   2
      Left            =   4200
      ScaleHeight     =   3915
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   10
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "âEÇ©ÇÁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Index           =   1
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   3960
      Width           =   4215
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "ëOÇ©ÇÁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Index           =   0
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "è„Ç©ÇÁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "1ñ{ñ⁄"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   7935
      Left            =   0
      Shape           =   2  'ë»â~
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p2(2), p3(1, 2), p4(999, 1, 2) As Integer
Public p4count As Integer
Dim pcol(1), mode As Integer

Private Sub Command1_Click()
For a = 0 To 5
Form1.List1(a).Clear
Next a
If p4count <> 0 Then
For a = 0 To p4count - 1
For b = 0 To 1
For c = 0 To 2
Form1.List1(b * 3 + c).AddItem (p4(a, b, c))
Next c
Next b
Next a
End If
End Sub

Private Sub Form_Load()
pcol(0) = RGB(0, 0, 255)
pcol(1) = RGB(255, 0, 0)
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
If mode = 0 Then
For a = 0 To 2
For b = 0 To 1
p3(b, a) = 0
Next b
Next a
For a = 0 To 2
p3(0, a) = p2(a)
Next a
ElseIf mode = 1 Then
For a = 0 To 2
If p3(0, a) = "?" Then p3(0, a) = p2(a)
Next a
End If
End If
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then
p2(0) = x
p2(1) = y
p2(2) = "?"
ElseIf Index = 1 Then
p2(0) = x
p2(2) = y
p2(1) = "?"
Else
p2(2) = x
p2(1) = y
p2(0) = "?"
End If
Call plas1
Call plas2
End Sub

Public Sub plas1()
For a = 0 To 2
If a = 0 Then
xx = p2(0)
yy = p2(1)
ElseIf a = 1 Then
xx = p2(0)
yy = p2(2)
ElseIf a = 2 Then
xx = p2(2)
yy = p2(1)
End If
Picture1(a).Cls
If xx = "?" Then
Picture1(a).Line (0, yy)-(Picture1(a).Width, yy), RGB(0, 255, 0)
ElseIf yy = "?" Then
Picture1(a).Line (xx, 0)-(xx, Picture1(a).Height), RGB(0, 255, 0)
Else
Picture1(a).Line (xx, yy - 200)-(xx, yy + 200), RGB(0, 255, 0)
Picture1(a).Line (xx - 200, yy)-(xx + 200, yy), RGB(0, 255, 0)
End If
Next a
Label1(0).Caption = "è„Ç©ÇÁ:x=" & p2(0) & ":y=" & p2(1)
Label1(2).Caption = "ëOÇ©ÇÁ:x=" & p2(0) & ":z=" & p2(2)
Label1(1).Caption = "âEÇ©ÇÁ:z=" & p2(2) & ":y=" & p2(1)
End Sub

Public Sub plas2()
For b = 0 To 1
For a = 0 To 2
If a = 0 Then
xx = p3(b, 0): If xx = "" Then xx = 0
yy = p3(b, 1): If yy = "" Then yy = 0
ElseIf a = 1 Then
xx = p3(b, 0): If xx = "" Then xx = 0
yy = p3(b, 2): If yy = "" Then yy = 0
ElseIf a = 2 Then
xx = p3(b, 2): If xx = "" Then xx = 0
yy = p3(b, 1): If yy = "" Then yy = 0
End If
If xx = "?" Then
Picture1(a).Line (0, yy)-(Picture1(a).Width, yy), pcol(b)
ElseIf yy = "?" Then
Picture1(a).Line (xx, 0)-(xx, Picture1(a).Height), pcol(b)
Else
Picture1(a).Line (xx, 0)-(xx, Picture1(a).Height), pcol(b)
Picture1(a).Line (0, yy)-(Picture1(a).Width, yy), pcol(b)
End If
Next a
Label2(0 + b * 3).Caption = "x=" & p3(b, 0) & ":y=" & p3(b, 1)
Label2(2 + b * 3).Caption = "x=" & p3(b, 0) & ":z=" & p3(b, 2)
Label2(1 + b * 3).Caption = "z=" & p3(b, 2) & ":y=" & p3(b, 1)
Next b
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
If mode = 0 Then
For a = 0 To 2
p3(1, a) = p2(a)
Next a
mode = 1
Else
For a = 0 To 2
If p3(1, a) = "?" Then p3(1, a) = p2(a)
Next a
For a = 0 To 2
For b = 0 To 1
If p3(b, a) = "?" Then c = 1
Next b
Next a
If c <> 1 Then
If Option1(0).Value = True Then
For a = 0 To 2
For b = 0 To 1
p4(p4count, b, a) = p3(b, a)
Next b
Next a
p4count = p4count + 1
Label3.Caption = p4count + 1 & "ñ{ñ⁄"
ElseIf Option1(1).Value = True Then
p4(p4count, 0, 0) = p3(0, 0)
p4(p4count, 0, 1) = p3(0, 1)
p4(p4count, 0, 2) = p3(0, 2)
p4(p4count, 1, 0) = p3(0, 0)
p4(p4count, 1, 1) = p3(0, 1)
p4(p4count, 1, 2) = p3(1, 2)
p4(p4count + 1, 0, 0) = p3(0, 0)
p4(p4count + 1, 0, 1) = p3(0, 1)
p4(p4count + 1, 0, 2) = p3(0, 2)
p4(p4count + 1, 1, 0) = p3(1, 0)
p4(p4count + 1, 1, 1) = p3(0, 1)
p4(p4count + 1, 1, 2) = p3(0, 2)
p4(p4count + 2, 0, 0) = p3(0, 0)
p4(p4count + 2, 0, 1) = p3(0, 1)
p4(p4count + 2, 0, 2) = p3(0, 2)
p4(p4count + 2, 1, 0) = p3(0, 0)
p4(p4count + 2, 1, 1) = p3(1, 1)
p4(p4count + 2, 1, 2) = p3(0, 2)
p4(p4count + 3, 0, 0) = p3(0, 0)
p4(p4count + 3, 0, 1) = p3(1, 1)
p4(p4count + 3, 0, 2) = p3(0, 2)
p4(p4count + 3, 1, 0) = p3(0, 0)
p4(p4count + 3, 1, 1) = p3(1, 1)
p4(p4count + 3, 1, 2) = p3(1, 2)
p4(p4count + 4, 0, 0) = p3(0, 0)
p4(p4count + 4, 0, 1) = p3(0, 1)
p4(p4count + 4, 0, 2) = p3(1, 2)
p4(p4count + 4, 1, 0) = p3(1, 0)
p4(p4count + 4, 1, 1) = p3(0, 1)
p4(p4count + 4, 1, 2) = p3(1, 2)
p4(p4count + 5, 0, 0) = p3(0, 0)
p4(p4count + 5, 0, 1) = p3(0, 1)
p4(p4count + 5, 0, 2) = p3(1, 2)
p4(p4count + 5, 1, 0) = p3(0, 0)
p4(p4count + 5, 1, 1) = p3(1, 1)
p4(p4count + 5, 1, 2) = p3(1, 2)
p4(p4count + 6, 0, 0) = p3(1, 0)
p4(p4count + 6, 0, 1) = p3(1, 1)
p4(p4count + 6, 0, 2) = p3(1, 2)
p4(p4count + 6, 1, 0) = p3(1, 0)
p4(p4count + 6, 1, 1) = p3(1, 1)
p4(p4count + 6, 1, 2) = p3(0, 2)
p4(p4count + 7, 0, 0) = p3(1, 0)
p4(p4count + 7, 0, 1) = p3(1, 1)
p4(p4count + 7, 0, 2) = p3(1, 2)
p4(p4count + 7, 1, 0) = p3(1, 0)
p4(p4count + 7, 1, 1) = p3(0, 1)
p4(p4count + 7, 1, 2) = p3(1, 2)
p4(p4count + 8, 0, 0) = p3(1, 0)
p4(p4count + 8, 0, 1) = p3(1, 1)
p4(p4count + 8, 0, 2) = p3(1, 2)
p4(p4count + 8, 1, 0) = p3(0, 0)
p4(p4count + 8, 1, 1) = p3(1, 1)
p4(p4count + 8, 1, 2) = p3(1, 2)
p4(p4count + 9, 0, 0) = p3(1, 0)
p4(p4count + 9, 0, 1) = p3(1, 1)
p4(p4count + 9, 0, 2) = p3(0, 2)
p4(p4count + 9, 1, 0) = p3(1, 0)
p4(p4count + 9, 1, 1) = p3(0, 1)
p4(p4count + 9, 1, 2) = p3(0, 2)
p4(p4count + 10, 0, 0) = p3(1, 0)
p4(p4count + 10, 0, 1) = p3(1, 1)
p4(p4count + 10, 0, 2) = p3(0, 2)
p4(p4count + 10, 1, 0) = p3(0, 0)
p4(p4count + 10, 1, 1) = p3(1, 1)
p4(p4count + 10, 1, 2) = p3(0, 2)
p4(p4count + 11, 0, 0) = p3(1, 0)
p4(p4count + 11, 0, 1) = p3(0, 1)
p4(p4count + 11, 0, 2) = p3(0, 2)
p4(p4count + 11, 1, 0) = p3(1, 0)
p4(p4count + 11, 1, 1) = p3(0, 1)
p4(p4count + 11, 1, 2) = p3(1, 2)
p4count = p4count + 12
Label3.Caption = p4count + 1 & "ñ{ñ⁄"
End If
End If
mode = 0
End If
End If
End Sub
