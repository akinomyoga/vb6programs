VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�P�ʂ̕ϊ�"
   ClientHeight    =   8190
   ClientLeft      =   4440
   ClientTop       =   1950
   ClientWidth     =   6045
   Icon            =   "�P�ʂ̕ϊ�.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "���ۑ�"
      Height          =   300
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.�P��1"
   End
   Begin VB.ListBox List2 
      Height          =   7080
      ItemData        =   "�P�ʂ̕ϊ�.frx":030A
      Left            =   2160
      List            =   "�P�ʂ̕ϊ�.frx":030C
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5895
   End
   Begin VB.ListBox List1 
      Height          =   7080
      ItemData        =   "�P�ʂ̕ϊ�.frx":030E
      Left            =   120
      List            =   "�P�ʂ̕ϊ�.frx":0310
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�E����
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "�ϊ���̐��l"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�ϊ��O�̒P��"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Menu mnfile 
      Caption         =   "�t�@�C��"
      Begin VB.Menu mntanifl 
         Caption         =   "�P�ʂ̏��̃t�@�C��[�P��1]���J��..."
      End
      Begin VB.Menu mntanifl2 
         Caption         =   "�P�ʂ̏��̃t�@�C��[�P��2]���J��..."
      End
      Begin VB.Menu mntanifl3 
         Caption         =   "�P�ʂ̏��̃t�@�C��[]���J��..."
      End
      Begin VB.Menu mnbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnn1 
         Caption         =   "�V�����P�ʂP�t�@�C��"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnn2 
         Caption         =   "�V�����P�ʂQ�t�@�C��"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnn3 
         Caption         =   "�V�����P�ʂQa�t�@�C��"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemcnt As Integer, bdat(999), cdat(2, 999) As Double
Dim �ő�P�ʐ� As Integer
Dim path1

Private Sub Command1_Click()
On Error GoTo errh
If List2.ListCount > 0 Then
CommonDialog1.Filter = "*.txt"
CommonDialog1.CancelError = True
CommonDialog1.ShowSave
b = CommonDialog1.FileName
If Right(b, 4) <> ".txt" Then b = b & ".txt"
Open b For Output As 1
Print #1, Text1.Text & List1.List(List1.ListIndex)
For a = 0 To List2.ListCount - 1
Print #1, "=" & List2.List(a)
Next a
Print #1, "*******����*******"
Close #1
End If
Exit Sub
errh:
If err.Number = 32755 Then
MsgBox "հ�ް�ɂ�跬ݾق���܂����B"
Else
MsgBox "�\�����ʴװ�F" & err.Number
On Error Resume Next
Close #1
End If
End Sub

Private Sub Form_Load()
�ő�P�ʐ� = 1000
path1 = Form1.path1
CommonDialog1.InitDir = path1
End Sub

Private Sub List1_Click()
Call Henkan
End Sub

Private Sub mnn1_Click()
Open path1 & "\�V.�P�ʂP" For Output As 1
Print #1, "�V�����P�ʂP�t�@�C���ł��B�P�ʂ�[�L��]�Ƃ��̒P�ʓ������[��"
Print #1, "����]���Ō��݂ɉ��s���Ȃ��珑�����݂܂��B�����炠��P�ʂP"
Print #1, "�t�@�C�����Q�l�ɂ��Ă��������B"
Close #1
End Sub

Private Sub mnn2_Click()
Open path1 & "\�V.�P�ʂQ" For Output As 1
Print #1, "�V�����P�ʂQ�t�@�C���ł��B�P�ʂ�[�L��]�Ƃ��̒P�ʓ�����̑傫"
Print #1, "���̔��[���q],[����],[���������_�̐����i������Ȃ��ꍇ��0"
Print #1, "�ɂ��Ă����΂悢�ł��B�j]�ŕ\�������A�i���ɒP�ʂQa�t�@�C����"
Print #1, "�Q�l�ɂ��镨������܂��B�����炠��P�ʂQ�t�@�C�����Q�l�ɂ���"
Print #1, "���������B�j����s��[,]�ŋ�؂��ď������݂܂��B"
Close #1
End Sub

Private Sub mnn3_Click()
Open path1 & "\�V.�P�ʂQa" For Output As 1
Print #1, "�V�����P�ʂQa�t�@�C���ł��B�P�ʂ̐ړ����[�L��]�Ƃ��̐ړ���"
Print #1, "������̑傫���̔��[���q],[����],[���������_�̐����i������"
Print #1, "�Ȃ��ꍇ��0�ɂ��Ă����΂悢�ł��B�j]�ŕ\����������s��[,]��"
Print #1, "��؂��ď������݂܂��B�����炠��P�ʂQa�t�@�C�����Q�l�ɂ���"
Print #1, "���������B"
Close #1
End Sub

Private Sub mntanifl_Click()
On Error GoTo errhand
CommonDialog1.Filter = "�P�ʂP�t�@�C��(*.�P�ʂP)|*.�P�ʂP"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
a = CommonDialog1.FileName
Dim d As Integer: d = 0
List1.Clear
Open a For Input As 1
Do While Not EOF(1) And d < �ő�P�ʐ�
Line Input #1, bb
Line Input #1, cc
bdat(d) = bb: cdat(0, d) = cc: List1.AddItem (bb)
d = d + 1
Loop
Close #1
itemcnt = d
Label3.Caption = a
Label4.Caption = itemcnt & "��"
Exit Sub
errhand:
If err.Number = 32755 Then
MsgBox "���[�U�[�ɂ��L�����Z������܂���!"
Else
MsgBox "�\�����ʴװ:" & err.nuber & "...�������t�@�C����I��ł��������I"
End If
End Sub

Private Sub mntanifl2_Click()
On Error GoTo errhand
CommonDialog1.Filter = "�P�ʂQ�t�@�C��(*.�P�ʂQ)|*.�P�ʂQ"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
a = CommonDialog1.FileName
adir = Left(a, Len(a) - Len(CommonDialog1.FileTitle))
Dim d As Integer
d = 0
List1.Clear
Open a For Input As 1 '1f
Do While Not EOF(1) And d < �ő�P�ʐ� '2f
Input #1, bb0, cc0, cc1, cc2, cc3, cc4
'�P�ʖ�,���q,����,�����_(e+/-?),�ړ���̧�ق̗L��(0/1),�����߽
If cc3 = 1 Then '3f
Dim edat(49), fdat(2, 49) As Double, d2 As Integer, cc2a As Integer
d2 = 0
Open adir & cc4 For Input As 2 '4f
Do While Not EOF(2) And d2 < 50 '5f
Input #2, ee0, ff0, ff1, ff2
'�ړ���,�ړ��̔{��(���q,����),���������_�̈ʒu
edat(d2) = ee0
fdat(0, d2) = ff0
fdat(1, d2) = ff1
fdat(2, d2) = ff2
d2 = d2 + 1
Loop '5e
Close #2 '4e
For cnt1 = 0 To d2 - 1 '4f
bdat(d) = edat(cnt1) & bb0
List1.AddItem (bdat(d))
cdat(0, d) = fdat(0, cnt1) * cc0
cdat(1, d) = fdat(1, cnt1) * cc1
cc2a = cc2
cdat(2, d) = fdat(2, cnt1) + cc2
d = d + 1
Next cnt1 '4e
Else '3m
bdat(d) = bb0
List1.AddItem (bdat(d))
cdat(0, d) = cc0
cdat(1, d) = cc1
cdat(2, d) = cc2
d = d + 1
End If '3e
Loop '2e
Close #1 '1e
itemcnt = d
Label3.Caption = a
Label4.Caption = itemcnt & "��"
Exit Sub
errhand:
If err.Number = 32755 Then
MsgBox "���[�U�[�ɂ��L�����Z������܂���!"
Else
MsgBox "�g���q���u�P�ʂQ�v�̃t�@�C����I��ł�������!�g���q���u�P�ʂQ�v�̃t�@�C����I��ł����̒��ӂ��\�������ꍇ���̃t�@�C���ɖ�肪����\��������܂��B"
MsgBox "�\�����ʴװ" & err.Number
End If
End Sub

Private Sub mntanifl3_Click()
'On Error GoTo errhand
CommonDialog1.Filter = "�P�ʃt�@�C��(*.units)|*.units"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
a = CommonDialog1.FileName
adir = Left(a, Len(a) - Len(CommonDialog1.FileTitle))
Dim d As Integer
d = 0
List1.Clear
Call read_file("", "", 1, 1, 0, adir, "\" & CommonDialog1.FileTitle, d) '�ďo
itemcnt = d
Label3.Caption = a
Label4.Caption = itemcnt & "��"
Exit Sub
errhand:
If err.Number = 32755 Then
MsgBox "���[�U�[�ɂ��L�����Z������܂���!"
Else
MsgBox "�������t�@�C���@���@���̊܂܂�Ă��Ȃ��t�@�C����I��ł�������!"
MsgBox "�\�����ʴװ" & err.Number
End If
End Sub

Public Sub read_file(ByVal b0, ByVal b1, ByVal a0, ByVal a1, ByVal a2, ByVal adir, ByVal adir2, ByRef d)
'>>>�p�X����>>>----------------------------------------------------------------------------
If Right(adir, 1) = "\" Then adir = Left(adir, Len(adir) - 1)
If Left(adir2, 1) = "." Then
 adir2 = Mid(adir2, 2)
 Do While Left(adir2, 1) = "."
  adir2 = Mid(adir2, 2)
  len0 = Len(adir)
  For n = len0 To 1
   If Mid(adir, n, 1) = "\" And Not tst = "tr" Then tst = "tr" & n1 = n - 1
  Next n
  If tst = "tr" Then adir = Left(adir, n1)
  tst = ""
Loop
End If
If Not Left(adir2, 1) = "\" Then adir2 = "\" & adir2
adir = adir & adir2
len0 = Len(adir)
For n = len0 To 1
 If Mid(adir, n, 1) = "\" And Not tst = "tr" Then tst = "tr" & n1 = n - 1
Next n
If tr = "tr" Then tr = "": adir0 = Left(adir, n1)
'>>>To Open The File>>>--------------------------------------------------------------------
flnm = FreeFile()
Open adir For Input As flnm
Do While Not EOF(flnm) And d < �ő�P�ʐ�
 Do
  Input #flnm, bb1
 Loop While Mid(bb1, 3, 1) = "!"
 Input #flnm, cc0, cc1, cc2
 '>>>����ؗ�����>>>------------------
 Do
  len0 = Len(bb1)
  For n = 1 To len0
   If Mid(bb1, n, 1) = ";" And Not tst = "tr" Then tst = "tr": n1 = n
  Next n
  If tst = "tr" Then
   tst = ""
   bb1 = Mid(bb0, n + 1)
   bb0 = Left(bb0, n - 1)
   tst2 = "come back"
  Else
   bb0 = bb1
   tst2 = ""
  End If
 '------------------<<<����ؗ�����<<<
  len0 = Len(bb0)
  For n = 1 To len0
   If Mid(bb0, n, 1) = "�e" And Not tst = "tr" Then tst = "tr": n1 = n + 1
  Next n
  If tst = "tr" Then
   '>>>�ڎ�����>>>-----------
   For n = len0 To n1
    If Mid(bb0, n, 1) = "�f" And Not tst = "trtr" Then tst = "trtr": n2 = n - n1
   Next n
   If tst = "trtr" Then
    bbx = Mid(bb0, n1, n2)
    bba = Left(bb0, n1 - 2)
    bbb = Mid(bb0, n1 + n2 + 1)
    Call read_file(b0 & bba, bbb & b1, a0 * cc0, a1 * cc1, a2 + cc2, adir0, bbx, d) '�ďo
   End If
   tst = ""
   '----------<<<�ڎ�����<<<
  Else
   '>>>�o�^>>>---------------
   bdat(d) = b0 & bb0 & b1
   List1.AddItem (bdat(d))
   cdat(0, d) = a0 * cc0
   cdat(1, d) = a1 * cc1
   cdat(2, d) = a2 + cc2
   MsgBox bdat(d) & " ### " & cdat(0, d) & " ### " & cdat(1, d) & " ### " & cdat(2, d)
   d = d + 1
   '---------------<<<�o�^<<<
  End If
 Loop While tst2 = "come back"
Loop
Close #flnm
End Sub

Private Sub Text1_Change()
Call Henkan
End Sub

Public Sub Henkan()
On Error GoTo errh
Dim c As Double, d As Double
If List1.ListIndex >= 0 Then '1f
b0 = cdat(0, List1.ListIndex) 'b0=�O��x/
b1 = cdat(1, List1.ListIndex): If b1 = 0 Then b1 = 1 'b1=�O��/x
b2 = cdat(2, List1.ListIndex) 'b2=�O�̏����_
On Error GoTo ERRH2: c = Text1.Text: On Error GoTo errh: Text1.Text = c 'c=�ϊ��O�̐��l
List2.Clear
For a = 0 To itemcnt - 1 '2f
Dim cda0 As Double, cda1 As Double
cda0 = cdat(0, a): If cda0 = 0 Then cda0 = 1 'cda0=���x/
cda1 = cdat(1, a): If cda1 = 0 Then cda1 = 1 'cda1=���/x
cda2 = cdat(2, a) 'cda2=��̏����_
ddd = c * b0 * cda1 / b1 / cda0 * (10 ^ (b2 - cda2))
List2.AddItem (ddd & " " & bdat(a))
'List2.AddItem (d / cda0 * cda1 e-cda2 & " " & bdat(a))
Next a '2e
End If '1e
ERRH2:
Exit Sub
errh:
MsgBox "error"
End Sub
