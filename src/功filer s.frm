VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'なし
   Caption         =   "Form2"
   ClientHeight    =   1545
   ClientLeft      =   4395
   ClientTop       =   3540
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  '砂時計
   ScaleHeight     =   1545
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "準備中・・・暫くお待ち下さい。"
      BeginProperty Font 
         Name            =   "HG正楷書体-PRO"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      MousePointer    =   11  '砂時計
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00008000&
      BorderStyle     =   1  '実線
      Caption         =   "ファイラ弌"
      BeginProperty Font 
         Name            =   "HG正楷書体-PRO"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      MousePointer    =   11  '砂時計
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ww = Form2.ScaleWidth - 15
hh = Form2.ScaleHeight - 15
Form2.Line (15, 15)-(ww, 15), RGB(255, 255, 255)
Form2.Line (15, 15)-(15, hh), RGB(255, 255, 255)
Form2.Line (0, hh)-(ww + 15, hh), 0
Form2.Line (ww, 0)-(ww, hh), 0
Form2.Line (15, hh - 15)-(ww, hh - 15), RGB(95, 95, 95)
Form2.Line (ww - 15, 15)-(ww - 15, hh - 15), RGB(95, 95, 95)
End Sub

Private Sub Timer1_Timer()
Form1.Show
Timer1.Interval = 0
End Sub
