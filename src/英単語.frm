VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ファイルを開く"
   ClientHeight    =   3825
   ClientLeft      =   1365
   ClientTop       =   3870
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "キャンセル"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2160
      Pattern         =   "*.功txt"
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2820
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin KBasic.ToggleButton ToggleButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "既存の物を削除"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub

Private Sub Command2_Click()
If Right(File1.Path, 1) = "\" Then
filepath = File1.Path & File1.FileName
Else
filepath = File1.Path & "\" & File1.FileName
End If
If ToggleButton1.Value = True Then Form1.List1.Clear
Form1.fileopen (filepath)
Call Command1_Click
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo errhand
Dir1.Path = Drive1.Drive
Exit Sub
errhand:
MsgBox "ドライブが見つかりません！"
End Sub

Private Sub Form_Load()
Dir1.Path = "."
End Sub

Private Sub ToggleButton1_Click()
If ToggleButton1.Value = False Then
ToggleButton1.Caption = "既存の物に追加"
Else
ToggleButton1.Caption = "既存の物を削除"
End If
End Sub
