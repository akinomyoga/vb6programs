VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "英単語"
   ClientHeight    =   3150
   ClientLeft      =   6195
   ClientTop       =   4155
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   3120
      ItemData        =   "英単語主.frx":0000
      Left            =   0
      List            =   "英単語主.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu mnfile 
      Caption         =   "ファイル"
      Begin VB.Menu mnopen 
         Caption         =   "開く..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mntouroku 
         Caption         =   "登録..."
         Shortcut        =   ^T
      End
      Begin VB.Menu save 
         Caption         =   "保存..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnkomoku 
      Caption         =   "項目"
      Begin VB.Menu mnkill 
         Caption         =   "選択されている項目を削除"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnup 
         Caption         =   "選択されている項目を上に移動"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mndown 
         Caption         =   "選択されている項目を下に移動"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub mndown_Click()
On Error GoTo errhand
a1 = List1.ListIndex
If a1 < List1.ListCount - 1 Then
a = List1.List(a1)
List1.RemoveItem List1.ListIndex
On Error GoTo err2
List1.AddItem a, (a1 + 1)
List1.ListIndex = a1 + 1
Else
MsgBox "これ以上、下へ移動できません。"
End If
Exit Sub
errhand:
MsgBox "項目が選択されていません！"
Exit Sub
err2:
MsgBox Err.Number
End Sub

Private Sub mnkill_Click()
On Error GoTo errhand
List1.RemoveItem List1.ListIndex
Exit Sub
errhand:
MsgBox "項目が選択されていません！"
End Sub

Private Sub mnopen_Click()
Form2.Show
End Sub

Private Sub mntouroku_Click()
Form3.Show
End Sub

Public Sub fileopen(filepath)
Open filepath For Input As #1
Do While Not EOF(1)
Input #1, Data1
Input #1, Data2
List1.AddItem Data1 & " " & Data2
Loop
Close #1
End Sub

Private Sub mnup_Click()
On Error GoTo errhand
a1 = List1.ListIndex
If a1 > 0 Then
a = List1.List(a1)
List1.RemoveItem List1.ListIndex
On Error GoTo err2
List1.AddItem a, (a1 - 1)
List1.ListIndex = a1 - 1
Else
MsgBox "これ以上、上へ移動できません。"
End If
Exit Sub
errhand:
MsgBox "項目が選択されていません！"
Exit Sub
err2:
MsgBox Err.Number
End Sub

Private Sub save_Click()
Form4.Show
End Sub

Public Sub filesave(filepath)
On Error GoTo errhand
Open filepath For Output As #1
For c = 0 To List1.ListCount - 1
Call word(List1.List(c))
Next c
Close #1
Exit Sub
errhand:
MsgBox "ちゃんと保存できませんでした！"
End Sub

Public Sub word(a)
Dim d As Boolean
d = False
For b = 1 To Len(a)
c = Mid(a, b, 1)
If c = " " Then
d = True
e = e & " " & f
f = ""
ElseIf d = True Then
f = f & c
Else
e = e & c
End If
Next b
Print #1, e
Print #1, f
End Sub

