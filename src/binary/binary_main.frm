VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Binary Viewer"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   574
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton bTab 
      Caption         =   "ログ"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   16
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton bTab 
      Caption         =   "数値"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton bTab 
      Caption         =   "画像"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton bTab 
      Caption         =   "file"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame TabPage 
      Caption         =   "binary データ"
      Height          =   5175
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton bRead 
         Caption         =   "読込(&R)"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Text            =   "Combo2"
         Top             =   720
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtBinDat 
         Height          =   4815
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"binary_main.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame TabPage 
      Height          =   5175
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton bDown 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton bRight 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton bCenter 
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton bUp 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox lstPixel 
         Height          =   300
         ItemData        =   "binary_main.frx":00BE
         Left            =   120
         List            =   "binary_main.frx":00D7
         TabIndex        =   23
         Text            =   "8"
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkInv 
         Caption         =   "左右反転"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton bLeft 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox lstBit 
         Height          =   300
         ItemData        =   "binary_main.frx":00F7
         Left            =   120
         List            =   "binary_main.frx":0113
         TabIndex        =   18
         Text            =   "2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton bStop 
         Caption         =   "停止(&S)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton bAnaly 
         Caption         =   "表示(&V)"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   1440
         ScaleHeight     =   321
         ScaleMode       =   3  'ﾋﾟｸｾﾙ
         ScaleWidth      =   457
         TabIndex        =   10
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lLst2 
         Caption         =   "px幅"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lLst1 
         Caption         =   "bit"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Frame TabPage 
      Caption         =   "読み込むファイルの設定"
      Height          =   5175
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   8415
      Begin VB.CommandButton bDir1 
         Caption         =   "デスクトップ"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "デスクトップ"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "binary_main.frx":013B
         Left            =   2880
         List            =   "binary_main.frx":0145
         TabIndex        =   5
         Text            =   "*.*"
         Top             =   240
         Width           =   5415
      End
      Begin VB.FileListBox File1 
         Height          =   4410
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   5415
      End
      Begin VB.DirListBox Dir1 
         Height          =   3870
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame TabPage 
      Caption         =   "ログ(処理の記録)"
      Height          =   5175
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   8415
      Begin RichTextLib.RichTextBox Log1 
         Height          =   4815
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8493
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"binary_main.frx":0160
      End
   End
   Begin VB.Label lFilename 
      BorderStyle     =   1  '実線
      Caption         =   "(ファイルはまだ選択されていません。)"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAnaly_Click()
message lFilename.Caption
If lFilename.Caption = "(ファイルはまだ選択されていません。)" Then
    message "ファイルが選択されていません。ファイルを選択してから再度実行してください。"
    Exit Sub
End If


bStop.Enabled = True

message openfile(lFilename.Caption)

Picture1.Cls
Picture1.Left = 1440
Select Case unitBit
Case 1
    Call readImage1bt
Case 2
    Call readImage2bt
Case 4
    Call readImage4bt
Case 24
    Call readImage24bt
Case 777
    Call readImagePoke
Case 778
    Call readImagePoke2
Case Else
    Call readImage
End Select

message closefile
End Sub

Private Sub bCenter_Click()
Picture1.Left = 1440
Picture1.Top = 240
End Sub

Private Sub bDir1_Click()
Dir1.Path = "C:\Documents and Settings\murase\デスクトップ"
End Sub

Private Sub bDown_Click()
Picture1.Top = Picture1.Top - 1350
End Sub

Private Sub bLeft_Click()
Picture1.Left = Picture1.Left + 1350
End Sub

Private Sub bRead_Click()
readStringH (lFilename.Caption)
End Sub

Private Sub bRight_Click()
Picture1.Left = Picture1.Left - 1350
End Sub

Private Sub bStop_Click()
bStop.Enabled = False
End Sub

Private Sub bTab_Click(index As Integer)
For Each i In TabPage()
 i.Visible = False
Next i
TabPage(index).Visible = True
End Sub

Private Sub bUp_Click()
Picture1.Top = Picture1.Top + 1350
End Sub

Private Sub Combo1_Change()
File1.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
File1.Pattern = Combo1.Text
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err1
Dir1.Path = Left(Drive1.Drive, 1) & ":\"
Exit Sub
err1:
Select Case Err.number
Case 68
MsgBox "ドライブが用意されていない可能性があります。今一度ドライブをご確認ください"
Drive1.Drive = Left(Dir1.Path, 2)
Case Else
MsgBox "予期せぬエラー - " & Err.number
End Select
End Sub

Private Sub File1_Click()
If File1.ListIndex >= 0 Then
path1 = File1.Path
If Right(path1, 1) <> "\" Then path1 = path1 + "\"
lFilename.Caption = path1 & File1.filename
End If
End Sub

Private Sub Form_Load()
Call bDir1_Click
End Sub

Public Sub message(x)
Select Case x
    Case "ファイルサイズ超過"
        Log1.Text = Log1.Text & "ファイルサイズが 1MB を超えています。実行に移すと処理しきれない恐れがありますので中止いたします。" & Chr(13)
    Case "成功"
        Log1.Text = Log1.Text & "ファイルの読み取りに無事に成功しました。" & Chr(13) & "****************************************************************" & Chr(13)
    Case Else
        Log1.Text = Log1.Text & x & Chr(13)
End Select
End Sub

Private Sub lstBit_Click()
Select Case lstBit.Text
Case "1", "2", "4", "8", "16", "24"
unitBit = lstBit.Text
unitBit2 = 2 ^ unitBit
Case "Poke縦"
unitBit = 777
Case "Poke横"
unitBit = 778
End Select
End Sub

Private Sub lstPixel_Click()
Select Case lstPixel.Text
Case "8", "16", "24", "32", "48", "64", "128"
pxWidth = lstPixel.Text
End Select
End Sub
