VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1620
   BeginProperty Font 
      Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   1620
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   600
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "‚Ü‚½—Ž‚Æ‚·"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "ƒoƒEƒ“ƒh.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double

Private Sub Command1_Click()
a = 0
Command1.Top = 0
End Sub

Private Sub Timer1_Timer()
a = a - 10
Command1.Top = Command1.Top - a
If Command1.Top > Form1.Height - 500 Then
a = -0.9 * a
Command1.Top = Form1.Height - 500
End If
End Sub
