VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   945
   ClientTop       =   1230
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13920
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   4800
      Picture         =   "功filer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "実行"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   3
      Left            =   4320
      Picture         =   "功filer.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   2
      Left            =   3960
      Picture         =   "功filer.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   1
      Left            =   3600
      Picture         =   "功filer.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   0
      Left            =   3240
      Picture         =   "功filer.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   0
      Picture         =   "功filer.frx":104A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   300
      Left            =   6480
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   8295
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   13231
            MinWidth        =   13231
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Bevel           =   0
            Object.Width           =   1773
            MinWidth        =   1773
            TextSave        =   "2019/07/10"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Bevel           =   0
            Object.Width           =   873
            MinWidth        =   873
            TextSave        =   "15:05"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5106
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "filename"
         Object.Tag             =   ""
         Text            =   "ﾌｧｲﾙ名"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "size"
         Object.Tag             =   ""
         Text            =   "ｻｲｽﾞ(ﾊﾞｲﾄ)"
         Object.Width           =   1483
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "更新日時"
         Object.Width           =   2542
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "読取専用"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "隠し"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ｱｰｶｲﾌﾞ"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ｼｽﾃﾑfile"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "拡張子"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "種類"
         Object.Width           =   2542
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9763
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4215
      Left            =   3240
      TabIndex        =   13
      Top             =   4080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"功filer.frx":138C
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   4455
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   7680
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":1424
            Key             =   "exe1"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":2076
            Key             =   "adobe1"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":2CC8
            Key             =   "rich1"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":391A
            Key             =   "lnk1"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":456C
            Key             =   "dll1"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":51BE
            Key             =   "some1"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":5E10
            Key             =   "press1"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":6662
            Key             =   "txt1"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":72B4
            Key             =   "w1"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":7F06
            Key             =   "hlp1"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":8B58
            Key             =   "ini1"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":97AA
            Key             =   "htm1"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":A3FC
            Key             =   "pict1"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":B04E
            Key             =   "pp1"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":BCA0
            Key             =   "of1"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":C8F2
            Key             =   "ol1"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":D544
            Key             =   "ax1"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":E196
            Key             =   "xl1"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   480
      Width           =   8655
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":EDE8
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":F33A
            Key             =   "adobe2"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":F68C
            Key             =   "lnk2"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":F9DE
            Key             =   "rich2"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":FD30
            Key             =   "press2"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":10282
            Key             =   "ini2"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":105D4
            Key             =   "htm2"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":10926
            Key             =   "hlp2"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":10C78
            Key             =   "of2"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":10FCA
            Key             =   "dll2"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":1131C
            Key             =   "ol2"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":1166E
            Key             =   "pp2"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":119C0
            Key             =   "xl2"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":11D12
            Key             =   "w2"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":12064
            Key             =   "ax2"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":123B6
            Key             =   "pict2"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":12708
            Key             =   "exe2"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":12A5A
            Key             =   "some2"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "功filer.frx":12DAC
            Key             =   "txt2"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function getParentPath(ByVal path As String) As String
    Index = InStrRev(path, "\")
    If Index = 0 Then
        getParentPath = ""
    Else
        getParentPath = Left(path, Index - 1)
    End If
End Function

Function endsWith(ByVal str As String, ByVal suffix As String) As Boolean
    endsWith = Right(str, Len(suffix)) = suffix
End Function

Function modifyDriveLetter(ByVal path As String) As String
    If Mid(path, 2, 2) = ":\" Then
        modifyDriveLetter = LCase(Left(path, 1)) & Right(path, Len(path) - 1)
    Else
        modifyDriveLetter = path
    End If
End Function

Sub recursiveExpand(ByVal path As String)
    Parent = getParentPath(path)
    If Parent = "" Then
        If endsWith(path, ":") Then
            Drive1.Drive = path
            path = LCase(path) & "\"
        End If
    Else
        recursiveExpand Parent
    End If
    TreeView1_Expand TreeView1.Nodes(path)
End Sub

Private Sub Command1_Click()
    desktopPath = modifyDriveLetter(CreateObject("WScript.Shell").SpecialFolders("Desktop"))
    recursiveExpand desktopPath
    TreeView1.SelectedItem = TreeView1.Nodes(desktopPath)
    Call TreeView1_NodeClick(TreeView1.Nodes(desktopPath))
End Sub

Private Sub Command2_Click(Index As Integer)
ListView1.View = Index
End Sub

Private Sub Command3_Click()
a = File1.path
If Right(a, 1) <> "\" Then a = a & "\"
c = ListView1.SelectedItem.Text
On Error Resume Next
b = Shell(a & c, 1)
If b = 0 Then
StatusBar1.Panels(1).Text = a & c & "の実行に失敗しました。"
Beep
Else
StatusBar1.Panels(1).Text = "タスクID" & b & "で" & a & c & "が実行されました。"
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo err:
a = Drive1.Drive
If Right(a, 1) <> "\" Then a = a & "\"
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , a, a, 1, 1
Call treerenew2(a)
Exit Sub
err:
Drive1.Drive = "c:"
End Sub

Private Sub File1_PathChange()
Dim b As ListItem
c = File1.path
ListView1.ListItems.Clear
If Right(c, 1) <> "\" Then c = c & "\"
For a = 0 To File1.ListCount - 1
Set b = ListView1.ListItems.Add(, File1.List(a), File1.List(a))
b.SubItems(1) = FileLen(c & File1.List(a))
b.SubItems(2) = FileDateTime(c & File1.List(a))
d = GetAttr(c & File1.List(a))
If d >= 32 Then
 d = d - 32: b.SubItems(5) = ""
End If
If d >= 4 Then
d = d - 4: b.SubItems(6) = ""
End If
If d >= 2 Then
d = d - 2: b.SubItems(4) = ""
End If
If d >= 1 Then b.SubItems(3) = ""
h = b.Text: i = Len(h)
Do
If e >= i Then
g = "": Exit Do
End If
g = f & g: f = Mid(h, i - e, 1): e = e + 1
Loop While f <> "."
b.SubItems(7) = g: e = 0: g = "": f = ""
j = filetype(b): b.SubItems(8) = j
Select Case j
Case "ｲﾒｰｼﾞ": b.Icon = "pict1": b.SmallIcon = "pict2"
Case "ﾀﾞｲﾅﾐｯｸﾘﾝｸﾗｲﾌﾞﾗﾘ": b.Icon = "dll1": b.SmallIcon = "dll2"
Case "ｱﾌﾟﾘｹｰｼｮﾝ": b.Icon = "exe1": b.SmallIcon = "exe2"
Case "ﾃｷｽﾄ文書": b.Icon = "txt1": b.SmallIcon = "txt2"
Case "MS Word": b.Icon = "w1": b.SmallIcon = "w2"
Case "MS Excel": b.Icon = "xl1": b.SmallIcon = "xl2"
Case "書庫": b.Icon = "press1": b.SmallIcon = "press2"
Case "MS Access": b.Icon = "ax1": b.SmallIcon = "ax2"
Case "MS OutlookExpress": b.Icon = "ol1": b.SmallIcon = "ol2"
Case "ﾍﾙﾌﾟﾌｧｲﾙ": b.Icon = "hlp1": b.SmallIcon = "hlp2"
Case "Webﾌｧｲﾙ": b.Icon = "htm1": b.SmallIcon = "htm2"
Case "MS PowerPoint": b.Icon = "pp1": b.SmallIcon = "pp2"
Case "MS Office": b.Icon = "of1": b.SmallIcon = "of2"
Case "設定ﾌｧｲﾙ": b.Icon = "ini1": b.SmallIcon = "ini2"
Case "Adobe Photoshop": b.Icon = "adobe1": b.SmallIcon = "adobe2"
Case "ﾘｯﾁﾃｷｽﾄ文書": b.Icon = "rich1": b.SmallIcon = "rich2"
Case "ｼｮｰﾄｶｯﾄ": b.Icon = "lnk1": b.SmallIcon = "lnk2"
Case Else: b.Icon = "some1": b.SmallIcon = "some2"
End Select
Next a
Label2.Caption = File1.path
End Sub

Private Sub Form_Load()
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "c:\", "c:\", 1, 1
Call treerenew2("c:\")
File1.path = "c:\"
Image1.Tag = Image1.Height
Form2.Hide
End Sub

Public Sub treerenew()

End Sub

Public Sub treerenew2(path1)
On Error GoTo ERR1
Dir1.path = path1
ll = Len(Dir1.List(-1))
If Right(path1, 1) <> "\" Then ll = ll + 1
If Dir1.ListCount > 0 Then
For a = 0 To Dir1.ListCount - 1
aa = Dir1.List(a)
TreeView1.Nodes.Add path1, 4, aa, Right(aa, Len(aa) - ll), 1, 1
Next a
End If
Exit Sub
ERR1:
Select Case err.Number
Case 71
MsgBox Left(Drive1.Drive, 1) & "ドライブは現在準備されていません!"
Drive1.Drive = "c:"
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
On Error GoTo ERR1
a = File1.path
If Right(a, 1) <> "\" Then a = a & "\"
b = Item.SubItems(8)
c = Item.SubItems(7)
If b = "ｲﾒｰｼﾞ" Then
Image1.Picture = LoadPicture(a & Item.Text)
Image1.ToolTipText = a & Item.Text
Image1.Stretch = False
If Image1.Height > Image1.Tag Then
Image1.Stretch = True
Image1.Height = Image1.Tag
End If
If Image1.Width > Image1.Tag Then
Image1.Stretch = True
Image1.Width = Image1.Tag
End If
End If
If b = "ﾃｷｽﾄ文書" Or b = "ﾘｯﾁﾃｷｽﾄ文書" Or b = "設定ﾌｧｲﾙ" Or c = "log" Or c = "LOG" Then
RichTextBox1.FileName = a & Item.Text
RichTextBox1.ToolTipText = a & Item.Text
End If
Exit Sub
ERR1:
StatusBar1.Panels(1) = "読み込めません!   エラーナンバー" & err.Number
Beep
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
aa = Node.Child.Index
For a = aa To aa + Node.Children - 1
If TreeView1.Nodes(a).Children = 0 Then
Call treerenew2(TreeView1.Nodes(a).Key)
End If
Next a
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
On Error GoTo ERR1
File1.path = Node.Key
Exit Sub
ERR1:
Select Case err.Number
Case 68
tex = Left(Drive1.Drive, 1) & "ドライブが準備されていない可能性があります。確かめて下さい。一旦Cドライブに戻します。"
MsgBox tex, , "デバイスの不備"
StatusBar1.Panels(1).Text = tex
Drive1.Drive = "c:"
End Select
End Sub

Public Function filetype(item1 As ListItem)
abc = oomojikomoji(item1.SubItems(7))
Select Case abc
Case "ico", "bmp", "dib", "wmf", "emf", "gif", "jpg", "jpeg", "jpe", "jfif", "cur", "sys"
filetype = "ｲﾒｰｼﾞ"
Case "dll": filetype = "ﾀﾞｲﾅﾐｯｸﾘﾝｸﾗｲﾌﾞﾗﾘ"
Case "exe": filetype = "ｱﾌﾟﾘｹｰｼｮﾝ"
Case "hlp", "chm", "col": filetype = "ﾍﾙﾌﾟﾌｧｲﾙ"
Case "lnk": filetype = "ｼｮｰﾄｶｯﾄ"
Case "txt", "dic", "exc", "scp"
filetype = "ﾃｷｽﾄ文書"
Case "doc", "dochtml", "docmhtml", "dot", "dothtml", "wiz", "wbk"
filetype = "MS Word"
Case "dif", "csv", "log", "slk", "xla", "xlb", "xlc", "xld", "xlk", "dqy", "iqy", "oqy"
filetype = "MS Excel"
Case "xll", "xlm", "xls", "xlshtml", "xlsmhtml", "xlt", "xlthtml", "xlv", "xlw", "rqy"
filetype = "MS Excel"
Case "arj", "bz2", "cab", "gz", "hqx", "lzh", "lzs", "mim", "zip", "rar", "tar", "taz", "tbz", "tgz", "uue", "xxe", "z"
filetype = "書庫"
Case "maq", "mar", "mas", "mat", "mau", "ade", "adn", "adp", "mad", "maf", "mag", "mam"
filetype = "MS Access"
Case "mda", "mdb", "mdbhtml", "mde", "mdn", "mdt", "mdw", "mdz", "ldb", "mav", "maw", "wizhtml"
filetype = "MS Access"
Case "fav", "eml", "hol", "msg", "nws", "oft"
filetype = "MS OutlookExpress"
Case "hta", "htm", "html", "mhtm", "mhtml", "url", "www", "wcs", "wsdl"
filetype = "Webﾌｧｲﾙ"
Case "pot", "pothtml", "ppa", "pps", "ppt", "ppthtml", "pptmhtml", "pwz"
filetype = "MS PowerPoint"
Case "acl", "det", "elm", "fad", "lex", "nick", "nk2", "obd", "obt", "obz", "odc", "oss", "ost"
filetype = "MS Office"
Case "pab", "pip", "prf", "pst", "rwz", "stf"
filetype = "MS Office"
Case "ini", "inf", "css", "scp"
filetype = "設定ﾌｧｲﾙ"
Case "8ba", "8bc", "8be", "8bf", "8bi", "8bp", "8bx", "8bs", "8by", "8li", "abr"
filetype = "Adobe Photoshop"
Case "acf", "aco", "act", "acv", "ado", "ahs", "ahu", "alv", "amp", "ams", "api"
filetype = "Adobe Photoshop"
Case "asp", "ast", "asv", "atf", "atn", "ava", "axt", "cha", "ffo", "grd", "psd", "psp"
filetype = "Adobe Photoshop"
Case "rtf", "dat", "mvd", "rpt", "bas", "yfs", "wri"
filetype = "ﾘｯﾁﾃｷｽﾄ文書"
End Select
End Function

Public Function oomojikomoji(abc)
For a = 1 To Len(abc)
a2 = Mid(abc, a, 1)
Select Case a2
Case "A": a1 = "a"
Case "B": a1 = "b"
Case "C": a1 = "c"
Case "D": a1 = "d"
Case "E": a1 = "e"
Case "F": a1 = "f"
Case "G": a1 = "g"
Case "H": a1 = "h"
Case "I": a1 = "i"
Case "J": a1 = "j"
Case "K": a1 = "k"
Case "L": a1 = "l"
Case "M": a1 = "m"
Case "N": a1 = "n"
Case "O": a1 = "o"
Case "P": a1 = "p"
Case "Q": a1 = "q"
Case "R": a1 = "r"
Case "S": a1 = "s"
Case "T": a1 = "t"
Case "U": a1 = "u"
Case "V": a1 = "v"
Case "W": a1 = "w"
Case "X": a1 = "x"
Case "Y": a1 = "y"
Case "Z": a1 = "z"
Case Else: a1 = a2
End Select
oomojikomoji = oomojikomoji & a1
Next a
End Function
