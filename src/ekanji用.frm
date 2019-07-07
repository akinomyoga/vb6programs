VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "eKanji óp"
   ClientHeight    =   870
   ClientLeft      =   9255
   ClientTop       =   7860
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "èîã¥ëÂäøòa"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ãûëÂçN‡Ü"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   15
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "5"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "uni"
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curd

Private Sub Command1_Click()
d = curd & "uni as " & Text1.Text & ".txt"
Open d For Output As 1
Print #1, "<SYSTEM>"
Print #1, "<Width = 64>"
Print #1, "<Height = 64>"
Print #1, "<Direction = 1>"
Print #1, "<TopSpace = 1>"
Print #1, "<BottomSpace = 1>"
Print #1, "<LeftSpace = 1>"
Print #1, "<RightSpace = 1>"
Print #1, "<CharacterPitch = 1>"
Print #1, "<LinePitch = 1>"
Print #1, "<Interlace = 1>"
Print #1, "</SYSTEM>"
Print #1, ""
Print #1, "<BODY>"
For x = 0 To 15
ProgressBar1.Value = x
xx = hhen(x)
For a = 0 To 15
aa = hhen(a)
For b = 0 To 15
bb = hhen(b)
c = c & "<u" & Text1.Text & xx & aa & bb & ">"
Next b
Next a
Print #1, c
c = ""
Next x
Print #1, "</BODY>"
Close #1
ProgressBar1.Value = 0
MsgBox "èIÇÌÇËÇ‹ÇµÇΩÅB"
End Sub

Private Sub Command2_Click()
d = curd & "ãûëÂçN‡Ü.txt"
Open d For Output As 1
Print #1, "<SYSTEM>"
Print #1, "<Width = 64>"
Print #1, "<Height = 64>"
Print #1, "<Direction = 1>"
Print #1, "<TopSpace = 1>"
Print #1, "<BottomSpace = 1>"
Print #1, "<LeftSpace = 1>"
Print #1, "<RightSpace = 1>"
Print #1, "<CharacterPitch = 1>"
Print #1, "<LinePitch = 1>"
Print #1, "<Interlace = 1>"
Print #1, "</SYSTEM>"
Print #1, ""
Print #1, "<BODY>"
For x = 0 To 12
ProgressBar1.Value = x * 15 / 12
For a = 1 To 4096
b = a + x * 4096
If b > 49188 Then Exit For
c = c & "<k" & b & ">"
Next a
Print #1, c
c = ""
Next x
Print #1, "</BODY>"
Close #1
ProgressBar1.Value = 0
MsgBox "èIÇÌÇËÇ‹ÇµÇΩÅB"
End Sub

Private Sub Command3_Click()
d = curd & "èîã¥ëÂäøòa.txt"
Open d For Output As 1
Print #1, "<SYSTEM>"
Print #1, "<Width = 64>"
Print #1, "<Height = 64>"
Print #1, "<Direction = 1>"
Print #1, "<TopSpace = 1>"
Print #1, "<BottomSpace = 1>"
Print #1, "<LeftSpace = 1>"
Print #1, "<RightSpace = 1>"
Print #1, "<CharacterPitch = 1>"
Print #1, "<LinePitch = 1>"
Print #1, "<Interlace = 1>"
Print #1, "</SYSTEM>"
Print #1, ""
Print #1, "<BODY>"
For x = 0 To 12
ProgressBar1.Value = x * 15 / 12
For a = 1 To 4096
b = a + x * 4096
If b > 50305 Then Exit For
c = c & "<m" & b & ">"
Next a
Print #1, c
c = ""
Next x
Print #1, "</BODY>"
Close #1
ProgressBar1.Value = 0
MsgBox "èIÇÌÇËÇ‹ÇµÇΩ"
End Sub

Private Sub Form_Load()
curd = CurDir
If Right(curd, 1) <> "\" Then curd = curd & "\"
End Sub

Public Function hhen(aa)
hhen = aa
Select Case aa
Case 10
hhen = "A"
Case 11
hhen = "B"
Case 12
hhen = "C"
Case 13
hhen = "D"
Case 14
hhen = "E"
Case 15
hhen = "F"
End Select
End Function
