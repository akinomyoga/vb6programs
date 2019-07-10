VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ý’è"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "F‚ÌÝ’è"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   2
         Left            =   1440
         Max             =   255
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   1
         Left            =   1440
         Max             =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   0
         Left            =   1440
         Max             =   255
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "ŽžŒv‚Q.frx":0000
         Left            =   120
         List            =   "ŽžŒv‚Q.frx":0019
         TabIndex        =   1
         Text            =   "•bj"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim whacol As Integer
Dim col(6, 2) As Integer
Private Sub Combo1_Change()
Select Case Combo1.Text
Case "•bj"
whacol = 2
Case "•ªj"
whacol = 1
Case "Žžj"
whacol = 0
Case "•bj‚Ì‹OÕ"
whacol = 3
Case "•ªj‚Ì‹OÕ"
whacol = 4
Case "Žžj‚Ì‹OÕ"
whacol = 5
Case "•¶Žš”Õ"
whacol = 6
Case Else
whacol = -1
End Select
If whacol >= 0 Then
For a = 0 To 2
HScroll1(a).Value = col(whacol, a)
Next a
End If
End Sub

Private Sub Combo1_Click()
Call Combo1_Change
End Sub

Private Sub Command1_Click()
For a = 0 To 5
For b = 0 To 2
Call Form1.colpass((a), (b), (col(a, b)))
Next b
Next a
Form1.Picture1.BackColor = RGB(col(6, 0), col(6, 1), col(6, 2))
Form1.ohr = -1
Form1.omr = -1
Form1.osr = -1
End Sub

Private Sub Form_Load()
Call Combo1_Change
End Sub

Private Sub HScroll1_Change(Index As Integer)
Dim b(2)
Label2(Index).Caption = HScroll1(Index).Value
If whacol >= 0 Then col(whacol, Index) = HScroll1(Index).Value
For a = 0 To 2
b(a) = HScroll1(a).Value
Next a
Label3.BackColor = RGB(b(0), b(1), b(2))
End Sub

Public Sub colpass(a As Integer, b As Integer, c As Integer)
col(a, b) = c
End Sub

Private Sub Text1_Change()
Select Case Text1.Text
Case "col14641"
For a = 0 To 6
Text1.Text = Text1.Text & Chr(13) & a & ":0=" & col(a, 0) & " " & a & ":1=" & col(a, 1) & " " & a & ":2=" & col(a, 2)
Next
MsgBox Text1.Text
Text1.Text = ""
Case "whacol14641"
Text1.Text = Text1.Text & Chr(13) & whacol
MsgBox Text1.Text
Text1.Text = ""
End Select
End Sub
