VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "プレイヤーの名前入力"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command3 
      Caption         =   "適応"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "HG正楷書体-PRO"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "HG正楷書体-PRO"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "プレイヤーの名前を入力して下さい。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Hide
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
End Sub

Private Sub Text1_Change(Index As Integer)
Command3.Enabled = True
End Sub
