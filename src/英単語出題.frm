VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�p�P��o��"
   ClientHeight    =   1380
   ClientLeft      =   3555
   ClientTop       =   4320
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "�t�@�C�����J��..."
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�I��"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   420
      ItemData        =   "�p�P��o��.frx":0000
      Left            =   0
      List            =   "�p�P��o��.frx":0002
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a1 As Long
Private Sub Command1_Click()
Label1.Caption = List1.List(a1)
a1 = a1 + 1
If a1 = List1.ListCount Then
a1 = 0
MsgBox "�܂��n�߂ɖ߂�܂�"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Public Sub fileopen(filepath)
Open filepath For Input As #1
Do While Not EOF(1)
Input #1, Data1
Input #1, Data2
List1.AddItem Data1
List1.AddItem Data2
Loop
Close #1
End Sub
