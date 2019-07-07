VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ŠG•`‚«"
   ClientHeight    =   4725
   ClientLeft      =   8520
   ClientTop       =   2325
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6630
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   2
      Top             =   4455
      Width           =   6375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4455
      Left            =   6375
      Max             =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ì×¯Ä
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   295
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   423
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu mnwin 
      Caption         =   "ƒEƒBƒ“ƒhƒE"
      Begin VB.Menu mnwintool 
         Caption         =   "“¹‹ï"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modes As Integer
Dim pdat(999, 999, 2) As Integer, adat(1), pmode(0)
Private Sub Form_Resize()
On Error GoTo err
If Form1.ScaleHeight < 1500 Then
Form1.Height = 1815
End If
If Form1.ScaleWidth < 3000 Then
Form1.Width = 1620
End If
HScroll1.Top = Form1.ScaleHeight - 255
VScroll1.Left = Form1.ScaleWidth - 255
HScroll1.Width = Form1.ScaleWidth - 255
VScroll1.Height = Form1.ScaleHeight - 255
Picture1.Height = HScroll1.Top
Picture1.Width = VScroll1.Left
Exit Sub
err:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnwintool_Click()
Form2.Show
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case pmode(0)
Case 0

End Select
End Sub
