VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows の既定値
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5415
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "aaa"
            TextSave        =   "aaa"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Animation1_Click()

End Sub

Private Sub Form_Load()
ProgressBar1.Value = 0
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox ProgressBar1.Value
End Sub

Private Sub Timer1_Timer()
On Error GoTo エラー
ProgressBar1.Value = ProgressBar1.Value + 1
StatusBar1.Panels(1).Text = ProgressBar1.Value
Exit Sub
エラー:
Timer1.Interval = 0
End Sub

