VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
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
      Height          =   3150
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
   Begin MSForms.ToggleButton ToggleButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;873"
      Value           =   "1"
      Caption         =   "既存の物を削除"
      FontName        =   "ＭＳ Ｐゴシック"
      FontHeight      =   165
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
filepath = File1.Path & File1.filename
Else
filepath = File1.Path & "\" & File1.filename
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
Dir1.Path = "c:\windows\ﾃﾞｽｸﾄｯﾌﾟ\"
End Sub

Private Sub ToggleButton1_Click()
If ToggleButton1.Value = False Then
ToggleButton1.Caption = "既存の物に追加"
Else
ToggleButton1.Caption = "既存の物を削除"
End If
End Sub
