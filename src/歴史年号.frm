VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "社会科歴史年号順整理Program"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "出来事"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "年号"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa(2050)

Private Sub Form_Load()
On Error GoTo err1
current1 = CurDir
If Right(current1, 1) <> "\" Then current1 = current1 & "\"
Open current1 & "年号.txt" For Input As 1
Do While Not EOF(1)
Input #1, a, b
Call TOUROKU(a, b)
Loop
Close #1
Exit Sub
err1:
MsgBox current1 & "年号.txtと言うファイルの読込に失敗しました。このファイルは、自分で作っていただいてかまいません。このプログラムはこれで終了します。"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err1
Open current1 & "年号.txt" For Output As 1
For a = 0 To 2051
If aa(a) <> "" Then
For b = 0 To Len(aa(a)) - 1
If Mid(aa(a), b, 1) <> " " Then
c = c & Mid(aa(a), b, 1)
Else
If c <> "" Then
Print #1, a & "," & c
c = ""
End If
End If
Next b
End If
Next a
Close #1
Exit Sub
err1:
MsgBox current1 & "年号.txtと言うファイルの書き込みに失敗しました。"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call TOUROKU(Text1.Text, Text2.Text)
Text1.Text = ""
Text2.Text = ""
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(KeyCode, Shift)
End Sub

Public Sub TOUROKU(tex1, tex2)
On Error GoTo err1
If tex1 > -1 And tex1 < 2051 Then
If tex2 <> "" Then
aa(tex1) = aa(tex1) & tex2 & " "
End If
Else
err1:
Label3.Caption = "年号の入力の仕方が間違っています！0から2050迄の数字を入力してください！"
End If
End Sub
