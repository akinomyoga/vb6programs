VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   ScaleHeight     =   675
   ScaleWidth      =   4965
   Tag             =   "0"
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label Label2 
         Alignment       =   2  'íÜâõëµÇ¶
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2175
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub Picture1_Click()

End Sub

Private Sub UserControl_EnterFocus()
Label1.Width = Picture1.Width / 100 * UserControl.Tag
Label1.Caption = UserControl.Tag + "Åì"
Label1.Caption = UserControl.Tag + "Åì"
End Sub

Private Sub UserControl_Initialize()
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
Label2.Width = UserControl.Width
Label1.Height = UserControl.Height
Label2.Height = UserControl.Height
End Sub
