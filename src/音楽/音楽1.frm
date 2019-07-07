VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "âπäyÇP"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   31
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   2040
   End
   Begin VB.CheckBox Check2 
      Caption         =   "âπÇÃì«çûèÄîı"
      BeginProperty Font 
         Name            =   "HGê≥û≤èëëÃ-PRO"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "âπÇÃì«çûê›íË"
      BeginProperty Font 
         Name            =   "HGê≥û≤èëëÃ-PRO"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CheckBox Check4 
      Caption         =   "èâä˙ê›íË"
      BeginProperty Font 
         Name            =   "HGê≥û≤èëëÃ-PRO"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "âπÇÃì«çû"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "âπäyver.ÇP"
      BeginProperty Font 
         Name            =   "HGä€∫ﬁºØ∏M-PRO"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Å¶ã÷â¸ïœ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Copy RightÅFë∫ê£"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 10
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Form1.Show
End Sub
