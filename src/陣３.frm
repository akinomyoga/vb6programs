VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command3 
      Caption         =   "ìKóp"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   4395
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      Left            =   2160
      Max             =   255
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   2160
      Max             =   255
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   2160
      Max             =   255
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "BLUE 0/255"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "GREEN 0/255"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "RED 0/255"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hide
Call Command3_Click
End Sub

Private Sub Command2_Click()
Hide
End Sub

Private Sub Command3_Click()
If Form1.trn2 = 1 Then
Form1.Picture1.BackColor = Picture1.BackColor
Form1.pl1cR = HScroll1.Value: Form1.pl1cG = HScroll2.Value: Form1.pl1cB = HScroll3.Value
Else
Form1.Picture2.BackColor = Picture1.BackColor
Form1.pl2cR = HScroll1.Value: Form1.pl2cG = HScroll2.Value: Form1.pl2cB = HScroll3.Value
End If
Call Form1.DRAWCELL
Call Form1.DRAWSTONE
Command3.Enabled = False
End Sub

Private Sub HScroll1_Change()
Call COLOR
Label1.Caption = "RED " & HScroll1.Value & "/255"
Command3.Enabled = True
End Sub

Private Sub HScroll2_Change()
Call COLOR
Label2.Caption = "GREEN " & HScroll2.Value & "/255"
Command3.Enabled = True
End Sub

Private Sub HScroll3_Change()
Call COLOR
Label3.Caption = "BLUE " & HScroll3.Value & "/255"
Command3.Enabled = True
End Sub

Public Sub COLOR()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
