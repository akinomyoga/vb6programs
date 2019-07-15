VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#9.0#0"; "KBasic.ocx"
Begin VB.Form Calc0 
   Caption         =   "ìdëÏ"
   ClientHeight    =   6780
   ClientLeft      =   3360
   ClientTop       =   2460
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8190
   Begin VB.CommandButton Command55 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   71
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command53 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   68
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   60
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command45 
      Caption         =   "ÅÄ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   59
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command44 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   58
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Å~"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   57
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Å{"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   56
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command33 
      Caption         =   "M1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   45
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command32 
      Caption         =   "M2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   44
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      Caption         =   "M3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   43
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command30 
      Caption         =   "M4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   42
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command29 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   40
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command26 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command25 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox TextY 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Text            =   "0"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox TextX 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Text            =   "0"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox TextM4 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Text            =   "0"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox TextM3 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Text            =   "0"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox TextM2 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox TextM1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Text            =   "0"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command21 
      Caption         =   "èIóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command17 
      Caption         =   "M4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "M3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      Caption         =   "M2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "M1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Text            =   "0"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Text            =   "0"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "<0>"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3600
      TabIndex        =   46
      Top             =   2520
      Width           =   4215
      Begin VB.CommandButton Command54 
         Caption         =   "-/+"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   70
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command52 
         Caption         =   "cot"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   67
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command51 
         Caption         =   "tan"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   66
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command50 
         Caption         =   "cosec"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   65
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command49 
         Caption         =   "cos"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   64
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command48 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   63
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command47 
         Caption         =   "sin"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   62
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command46 
         Caption         =   "ÉŒ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   61
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command41 
         Caption         =   "YèÊç™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   54
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command40 
         Caption         =   "XèÊç™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command39 
         Caption         =   "óßï˚ç™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   52
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command38 
         Caption         =   "ïΩï˚ç™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         Caption         =   "YèÊ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   50
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command36 
         Caption         =   "XèÊ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   49
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command35 
         Caption         =   "óßï˚"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   48
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command34 
         Caption         =   "ïΩï˚"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin KBasic.ToggleButton ToggleButton1 
         Height          =   375
         Left            =   1560
         TabIndex        =   69
         Top             =   1800
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "êîéöÇécÇ∑"
         Value           =   -1  'True
      End
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   840
      TabIndex        =   55
      Top             =   6240
      Width           =   5655
   End
   Begin VB.Label Label6 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "M4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "M3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "M2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "M1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   375
   End
End
Attribute VB_Name = "Calc0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M

Private Sub Command1_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 1
Else
Text1.Text = Text1.Text & 1
End If
End Sub

Private Sub Command10_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 0
Else
Text1.Text = Text1.Text & 0
End If
End Sub

Private Sub Command11_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 0
Else
Text1.Text = Text1.Text & 0
Text1.Text = Text1.Text & 0
End If
End Sub

Private Sub Command12_Click()
Text1.Text = 0
Text2.Text = 0
TextM1.Text = 0
TextM2.Text = 0
TextM3.Text = 0
TextM4.Text = 0
TextX.Text = 0
TextY.Text = 0
End Sub

Private Sub Command13_Click()
M = Text1.Text
Text1.Text = 0
End Sub

Private Sub Command14_Click()
a = TextM1.Text
TextM1.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command15_Click()
a = TextM2.Text
TextM2.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command16_Click()
a = TextM3.Text
TextM3.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command17_Click()
a = TextM4.Text
TextM4.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command18_Click()
a = TextX.Text
TextX.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command19_Click()
a = TextY.Text
TextY.Text = Text2.Text
Text2.Text = a
End Sub

Private Sub Command2_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 2
Else
Text1.Text = Text1.Text & 2
End If
End Sub

Private Sub Command20_Click()
Text2.Text = 0
End Sub

Private Sub Command21_Click()
End
End Sub

Private Sub Command22_Click()
TextM1.Text = 0
End Sub

Private Sub Command23_Click()
TextM2.Text = 0
End Sub

Private Sub Command24_Click()
TextM3.Text = 0
End Sub

Private Sub Command25_Click()
TextM4.Text = 0
End Sub

Private Sub Command26_Click()
TextX.Text = 0
End Sub

Private Sub Command27_Click()
TextY.Text = 0
End Sub

Private Sub Command28_Click()
a = TextY.Text
TextY.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command29_Click()
a = TextX.Text
TextX.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command3_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 3
Else
Text1.Text = Text1.Text & 3
End If
End Sub

Private Sub Command30_Click()
a = TextM4.Text
TextM4.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command31_Click()
a = TextM3.Text
TextM3.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command32_Click()
a = TextM2.Text
TextM2.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command33_Click()
a = TextM1.Text
TextM1.Text = Text1.Text
Text1.Text = a
End Sub

Private Sub Command34_Click()
Text1.Text = Text1.Text ^ 2
End Sub

Private Sub Command35_Click()
Text1.Text = Text1.Text ^ 3
End Sub

Private Sub Command36_Click()
Text1.Text = Text1.Text ^ TextX.Text
End Sub

Private Sub Command37_Click()
Text1.Text = Text1.Text ^ TextY.Text
End Sub

Private Sub Command38_Click()
Text1.Text = Text1.Text ^ (1 / 2)
End Sub

Private Sub Command39_Click()
Text1.Text = Text1.Text ^ (1 / 3)
End Sub

Private Sub Command4_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 4
Else
Text1.Text = Text1.Text & 4
End If
End Sub

Private Sub Command40_Click()
If x <> 0 Then
Text1.Text = Text1.Text ^ (1 / TextX.Text)
Else
Label7.Caption = "0èÊç™ÇÕèoóàÇ‹ÇπÇÒÅI"
End If
End Sub

Private Sub Command41_Click()
If y <> 0 Then
Text1.Text = Text1.Text ^ (1 / TextY.Text)
Else
Label7.Caption = "0èÊç™ÇÕèoóàÇ‹ÇπÇÒÅI"
End If
End Sub

Private Sub Command42_Click()
Dim a As Single
Dim b As Single
Dim c As Single
On Error GoTo MSG
a = Text1.Text
b = Text2.Text
c = a + b
Text2.Text = c
If ToggleButton1.Value = True Then
Else
Text1.Text = 0
End If
Exit Sub
MSG:
Label7.Caption = "è¨êîì_ÇÕÅAÇQå¬à»è„ïtÇØÇÁÇÍÇ‹ÇπÇÒÅI"
End Sub

Private Sub Command43_Click()
On Error GoTo MSG
Text2.Text = Text1.Text * Text2.Text
If ToggleButton1.Value = True Then
Else
Text1.Text = 0
End If
Exit Sub
MSG:
Label7.Caption = "è¨êîì_ÇÕÅAÇQå¬à»è„ïtÇØÇÁÇÍÇ‹ÇπÇÒÅI"
End Sub

Private Sub Command44_Click()
On Error GoTo MSG
Text2.Text = Text2.Text - Text1.Text
If ToggleButton1.Value = True Then
Else
Text1.Text = 0
End If
Exit Sub
MSG:
Label7.Caption = "è¨êîì_ÇÕÅAÇQå¬à»è„ïtÇØÇÁÇÍÇ‹ÇπÇÒÅI"
End Sub

Private Sub Command45_Click()
On Error GoTo MSG
Text2.Text = Text2.Text / Text1.Text
If ToggleButton1.Value = True Then
Else
Text1.Text = 0
End If
Exit Sub
MSG:
Label7.Caption = "è¨êîì_ÇÕÅAÇQå¬à»è„ïtÇØÇÁÇÍÇ‹ÇπÇÒÅI"
End Sub

Private Sub Command46_Click()
If Text1.Text = 0 Then
Text1.Text = 3.14159265358979
Else
Text1.Text = Text1.Text * 3.14159265358979
End If
End Sub

Private Sub Command47_Click()
Text1.Text = Sin(Text1.Text)
End Sub

Private Sub Command48_Click()
Text1.Text = 1 / Sin(Text1.Text)
End Sub

Private Sub Command49_Click()
Text1.Text = Cos(Text1.Text)
End Sub

Private Sub Command5_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 5
Else
Text1.Text = Text1.Text & 5
End If
End Sub

Private Sub Command50_Click()
Text1.Text = 1 / Cos(Text1.Text)
End Sub

Private Sub Command51_Click()
Text1.Text = Tan(Text1.Text)
End Sub

Private Sub Command52_Click()
Text1.Text = 1 / Tan(Text1.Text)
End Sub

Private Sub Command53_Click()
Text1.Text = M
End Sub

Private Sub Command54_Click()
On Error GoTo MSG
Text1.Text = 0 - Text1.Text
Exit Sub
MSG:
Label7.Caption = "è¨êîì_ÇÕÅAÇQå¬à»è„ïtÇØÇÁÇÍÇ‹ÇπÇÒÅI"
End Sub

Private Sub Command55_Click()
M = Text1.Text
Text1.Text = Text1.Text & "."
End Sub

Private Sub Command6_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 6
Else
Text1.Text = Text1.Text & 6
End If
End Sub

Private Sub Command7_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 7
Else
Text1.Text = Text1.Text & 7
End If
End Sub

Private Sub Command8_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 8
Else
Text1.Text = Text1.Text & 8
End If
End Sub

Private Sub Command9_Click()
M = Text1.Text
If Text1.Text = "0" Then
Text1.Text = 9
Else
Text1.Text = Text1.Text & 9
End If
End Sub

Private Sub Form_Activate()
M = 0
End Sub

Private Sub Text1_Change()
Frame1.Caption = "<" & Text1.Text & ">"
End Sub


