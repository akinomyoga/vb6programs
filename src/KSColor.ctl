VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Begin VB.UserControl KSColor 
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   1320
   ScaleWidth      =   3495
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2760
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2760
      TabIndex        =   2
      Text            =   "0"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2760
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin KBasic.SpinButton SpinButton3 
      Height          =   255
      Left            =   3120
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Max             =   255
      ForeColor       =   16711680
   End
   Begin KBasic.SpinButton SpinButton2 
      Height          =   255
      Left            =   3120
      Top             =   360
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Max             =   255
      ForeColor       =   65280
   End
   Begin KBasic.SpinButton SpinButton1 
      Height          =   255
      Left            =   3120
      Top             =   120
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Max             =   255
      ForeColor       =   255
   End
   Begin VB.HScrollBar ScrollBar3 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.HScrollBar ScrollBar2 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.HScrollBar ScrollBar1 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "000000"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "KScolor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private Sub Picture1_Change()
    Label1.Caption = Hex$(scroll1bar.Value) & Hex$(ScrollBar2.Value) & Hex$(ScrollBar3.Value)
End Sub

Private Sub ScrollBar1_Change()
    Picture1.BackColor = RGB(ScrollBar1.Value, ScrollBar2.Value, ScrollBar3.Value)
    Text1.Text = ScrollBar1.Value
    SpinButton1.Value = ScrollBar1.Value
End Sub

Private Sub ScrollBar2_Change()
    Picture1.BackColor = RGB(ScrollBar1.Value, ScrollBar2.Value, ScrollBar3.Value)
    Text2.Text = ScrollBar2.Value
    SpinButton2.Value = ScrollBar2.Value
End Sub

Private Sub ScrollBar3_Change()
    Picture1.BackColor = RGB(ScrollBar1.Value, ScrollBar2.Value, ScrollBar3.Value)
    Text3.Text = ScrollBar3.Value
    SpinButton3.Value = ScrollBar3.Value
End Sub

Private Sub SpinButton1_Change()
    ScrollBar1.Value = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    ScrollBar2.Value = SpinButton2.Value
End Sub

Private Sub SpinButton3_Change()
    ScrollBar3.Value = SpinButton3.Value
End Sub
