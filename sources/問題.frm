VERSION 5.00
Begin VB.Form 問題 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command3 
      Caption         =   "終了"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開く"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "HG正楷書体-PRO"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "問題"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q(1 To 100)
Dim a(1 To 100)
Dim l As Integer
Dim m As Integer
Dim n As Integer


Private Sub Command1_Click()
On Error GoTo NASI
Open App.Path & Text1.Text & ".txt" For Input As #1
For b = 1 To 100
Line Input #1, q(b)
Line Input #1, a(b)
Next b
Close #1
Label2.Caption = "あと100問"
Label3.Caption = "0問正解"
l = 0
m = 100
n = 0
MsgBox "無事にファイルを開くことが出来ました。"
Exit Sub
NASI:
MsgBox "指定したファイルがないか、内容の正しくないファイルです。"
End Sub

Private Sub Command2_Click()
If Text2.Text = a(l) Then
n = n + 1: Label3.Caption = n & "問正解"
MsgBox "正解！"
Else
MsgBox "残念、不正解！正しい答えは「　" & a(l) & "　」だよ。"
End If
m = m - 1: Label1.Caption = "あと" & m & "問"
l = l + 1: If l = 101 Then GoTo OWARI
On Error GoTo DEKIN
Label1.Caption = q(l)
Text2.Text = ""
Exit Sub
OWARI:
Label1.Caption = n & "点!!!"
MsgBox "もう終わりです。あなたの得点は、" & Label1.Caption & "です。"
Exit Sub
DEKIN:
MsgBox "ファイルが指定されていないか、もう終わってしまったかで出来ません。"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
l = 0
m = 0
n = 0
End Sub
