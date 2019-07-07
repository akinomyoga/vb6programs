VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   Caption         =   "âpíPåÍÇÃìoò^"
   ClientHeight    =   1965
   ClientLeft      =   1365
   ClientTop       =   1215
   ClientWidth     =   4605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "ìoò^èIÇÌÇË"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ìoò^"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "âpíPåÍÇÃà”ñ°ì¸óÕ"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "âpíPåÍÅ@ì¸óÕ"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" And Text2.Text <> "" Then
Form1.List1.AddItem Text1.Text & " " & Text2.Text
Text1.Text = ""
Text2.Text = ""
Else
MsgBox "ÇøÇ·ÇÒÇ∆ì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅI"
End If
End Sub

Private Sub Command2_Click()
Form3.Hide
End Sub
