VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "英単語保存"
   ClientHeight    =   2370
   ClientLeft      =   6195
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "キャンセル"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ファイル名を入力して下さい！"
Else
If Right(Dir1.Path, 1) = "\" Then
filepath = Dir1.Path & Text1.Text & ".功txt"
Else
filepath = Dir1.Path & "\" & Text1.Text & ".功txt"
End If
Form1.filesave (filepath)
MsgBox filepath & "に保存されました！"
Form4.Hide
End If
End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Drive1_Change()
On Error GoTo errhand
Dir1.Path = Drive1.Drive
Exit Sub
errhand:
MsgBox "ドライブが見つかりません！"
End Sub

Private Sub Form_Load()
Dir1.Path = "c:\windows\ﾃﾞｽｸﾄｯﾌﾟ\"
End Sub
