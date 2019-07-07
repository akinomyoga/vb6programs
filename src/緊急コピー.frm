VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "緊急コピー"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command4 
      Caption         =   "kill file"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "copy to a:(FD)"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   720
      TabIndex        =   4
      Text            =   "aaa"
      Top             =   3270
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "make                       .zip"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "copy to A:(FD)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2610
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo erred
b = File1.Path
If Right(b, 1) <> "\" Then b = b & "\"
For a = 1 To File1.ListCount
c = File1.List(a - 1)
FileCopy b & c, "a:\" & c
Next a
Exit Sub
erred:
MsgBox "err"
End Sub

Private Sub Command2_Click()
On Error GoTo erred
a = File1.Path
If Right(a, 1) <> "\" Then a = a & "\"
FileCopy "c:\windows\ﾃﾞｽｸﾄｯﾌﾟ\aaa.txt", a & Text1.Text & ".zip"
Exit Sub
erred:
MsgBox "err"
End Sub

Private Sub Command3_Click()
On Error GoTo erred
If File1.ListIndex = -1 Then Exit Sub
a = File1.Path
If Right(a, 1) <> "\" Then a = a & "\"
FileCopy a & File1.List(File1.ListIndex), "a:\" & File1.List(File1.ListIndex)
Exit Sub
erred:
MsgBox "err"
End Sub

Private Sub Command4_Click()
On Error GoTo erred
a = MsgBox("ファイルが削除されますよろしいですか？", vbOKCancel)
If a <> 1 Then Exit Sub
a = File1.Path
If Right(a, 1) <> "\" Then a = a & "\"
Kill a & File1.List(File1.ListIndex)
Exit Sub
erred:
MsgBox "err"
End Sub

Private Sub Dir1_Change()
On Error GoTo erred
File1.Path = Dir1.Path
Exit Sub
erred:
MsgBox "err"
End Sub

Private Sub File1_Click()
On Error GoTo erred
If File1.ListIndex = -1 Then Exit Sub
a = File1.Path
If Right(a, 1) <> "\" Then a = a & "\"
b = FileLen(a & File1.List(File1.ListIndex))
Label1.Caption = b & " バイト"
Exit Sub
erred:
MsgBox "err"
End Sub
