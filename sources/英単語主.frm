VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�p�P��"
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
      ItemData        =   "�p�P���.frx":0000
      Left            =   0
      List            =   "�p�P���.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu mnfile 
      Caption         =   "�t�@�C��"
      Begin VB.Menu mnopen 
         Caption         =   "�J��..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mntouroku 
         Caption         =   "�o�^..."
         Shortcut        =   ^T
      End
      Begin VB.Menu save 
         Caption         =   "�ۑ�..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnkomoku 
      Caption         =   "����"
      Begin VB.Menu mnkill 
         Caption         =   "�I������Ă��鍀�ڂ��폜"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnup 
         Caption         =   "�I������Ă��鍀�ڂ���Ɉړ�"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mndown 
         Caption         =   "�I������Ă��鍀�ڂ����Ɉړ�"
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
MsgBox "����ȏ�A���ֈړ��ł��܂���B"
End If
Exit Sub
errhand:
MsgBox "���ڂ��I������Ă��܂���I"
Exit Sub
err2:
MsgBox Err.Number
End Sub

Private Sub mnkill_Click()
On Error GoTo errhand
List1.RemoveItem List1.ListIndex
Exit Sub
errhand:
MsgBox "���ڂ��I������Ă��܂���I"
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
MsgBox "����ȏ�A��ֈړ��ł��܂���B"
End If
Exit Sub
errhand:
MsgBox "���ڂ��I������Ă��܂���I"
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
MsgBox "�����ƕۑ��ł��܂���ł����I"
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

