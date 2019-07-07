VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "拡張子検索"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   4200
      ScaleHeight     =   3015
      ScaleWidth      =   3615
      TabIndex        =   11
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command6 
         Caption         =   "読込"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "消去"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "並替"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   2370
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "検索.frx":0000
      Left            =   120
      List            =   "検索.frx":000A
      TabIndex        =   9
      Text            =   "拡張子"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4230
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   3975
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   0
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Height          =   2610
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.FileListBox File1 
         Height          =   2400
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "検索するファイル"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検索"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------------
Dim canceled As Boolean
Dim dat1(999)

Private Sub Command1_Click()
    canceled = False
    Form1.Width = Picture2.Width + 240
    Picture1.Visible = False
    Picture2.Move 120, 120
    Command3.Enabled = False
    Command2.Enabled = True
    Command1.Enabled = False
    Select Case Combo1.Text
    Case "拡張子": Call selectfold
    Case "拡張子を持つファイル": Call selectfold2
    End Select
    Call Command2_Click
End Sub

Private Sub Command2_Click()
    Picture2.Move Picture1.Width + 240, 120
    Form1.Width = Picture2.Width + Picture2.Left + 120
    Picture1.Visible = True
    Command3.Enabled = True
    Command2.Enabled = False
    Command1.Enabled = True
    canceled = True
    MsgBox "検索を終了します。"
    StatusBar1.Panels(1).Text = "検索終了:" & List1.ListCount & "項目検出"
End Sub

Private Sub Command3_Click() '情報を保存
    lc = List1.ListCount - 1
    If lc < 0 Then Exit Sub
    Open App.Path & "\" & Combo1.Text & ".txt" For Output As 1
    For ls = 0 To lc
        Print #1, List1.List(ls)
    Next ls
    Close #1
End Sub

Private Sub Command5_Click() '情報を初期化
    List1.Clear
End Sub

Private Sub Command6_Click() '保存しておいた情報を読み込む
    On Error GoTo err1
    Open App.Path & "\" & Combo1.Text & ".txt" For Input As 1
    'n = 0
    Do While Not EOF(1) And n < 1000
        Line Input #1, a
        m = List1.ListCount - 1
        If m < 0 Then GoTo skp1
        For b = 0 To m
            If List1.List(b) = a Then GoTo skp2
        Next b
skp1:         List1.AddItem a
skp2:
        'dat1(n) = a
        'n = n + 1
    Loop
    Close #1
    Exit Sub
err1:
    Select Case Err.Number
    Case 53: MsgBox "ファイルがありません"
    Case Else: MsgBox "何らかのエラー"
    End Select
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Picture1.Move 120, 120
    Picture2.Move Picture1.Width + 240, 120
    Form1.Width = Picture2.Width + Picture2.Left + 120
    Picture1.Visible = True
    Exit Sub '##################################
    Open App.Path & "\拡張子.txt" For Input As 1
    n = 0
    Do While Not EOF(1) And n < 1000
        Line Input #1, a
        dat1(n) = a
        n = n + 1
    Loop
    Close #1
End Sub

Public Sub selectfold() '使われている拡張子
    Path = Dir1.Path
    StatusBar1.Panels(1).Text = Path & "を検索中"
    File1.Path = Path
    m2 = File1.ListCount - 1
    If m2 >= 0 Then
        For ic = 0 To m2
            a = Split(File1.List(ic), ".")
            If UBound(a) > 0 Then
                itm3 = LCase(a(UBound(a)))
                m1 = List1.ListCount - 1
                If m1 < 0 Then GoTo skp1
                For ic2 = 0 To m1
                    If List1.List(ic2) = itm3 Then GoTo skp2
                Next ic2
skp1:
                List1.AddItem (itm3)
skp2:
            End If
        Next ic
    End If
    b = Dir1.ListCount - 1
    If b = 0 Then Exit Sub
    For a = 0 To b
        Dir1.Path = Dir1.List(a)
        Call Form1.selectfold
        If canceled Then Exit Sub
        Dir1.Path = Path
        File1.Path = Path
    Next a
    DoEvents
    If cancelsd Then Exit Sub
End Sub

Public Sub selectfold2() '拡張子のあるファイル
    Path = Dir1.Path
    StatusBar1.Panels(1).Text = Path & "を検索中"
    File1.Path = Path
    m2 = File1.ListCount - 1
    If m2 >= 0 Then
        For ic = 0 To m2
            a = Split(File1.List(ic), ".")
            If UBound(a) > 0 Then
                itm3 = LCase(a(UBound(a)))
                m1 = List1.ListCount - 1
                If itm3 = Text1.Text Then List1.AddItem (Path & "\" & File1.List(ic))
            End If
        Next ic
    End If
    b = Dir1.ListCount - 1
    If b = 0 Then Exit Sub
    For a = 0 To b
        Dir1.Path = Dir1.List(a)
        Call Form1.selectfold2
        If canceled Then Exit Sub
        Dir1.Path = Path
        File1.Path = Path
    Next a
    DoEvents
    If cancelsd Then Exit Sub
End Sub

Private Sub List1_Click()
    MsgBox List1.List(List1.ListIndex)
End Sub
