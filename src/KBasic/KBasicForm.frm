VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form KBasicForm 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin KBasic.ScrollBar ScrollBar1 
      Height          =   135
      Left            =   120
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin KBasic.ScrollBar ScrollBar3 
      Height          =   255
      Left            =   120
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Max             =   5
      BackColor       =   12632319
   End
   Begin KBasic.SpinButton SpinButton5 
      Height          =   375
      Left            =   3000
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
   End
   Begin KBasic.ToggleButton ToggleButton3 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Value           =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin KBasic.SpinButton SpinButton2 
      Height          =   255
      Left            =   480
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Orientation     =   1
   End
   Begin KBasic.SpinButton SpinButton4 
      Height          =   375
      Left            =   1920
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Orientation     =   1
      BackColor       =   12648384
      ForeColor       =   32768
   End
   Begin KBasic.SpinButton SpinButton1 
      Height          =   255
      Left            =   120
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin KBasic.ColorButton ColorButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Color1"
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KBasic.ToggleButton ToggleButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Toggle2"
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KBasic.ToggleButton ToggleButton1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Toggle1"
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "XYZ"
      BackColor       =   12648447
      ForeColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin KBasic.SpinButton SpinButton3 
      Height          =   375
      Left            =   1560
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColor       =   12632319
      ForeColor       =   128
   End
   Begin KBasic.ColorButton ColorButton2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Color1"
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   16761024
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KBasic.ColorButton ColorButton3 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Color1"
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BackColor       =   16761024
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KBasic.SpinButton SpinButton6 
      Height          =   375
      Left            =   3360
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Orientation     =   1
      BackColor       =   12648384
      ForeColor       =   32768
      Enabled         =   0   'False
   End
   Begin KBasic.ToggleButton ToggleButton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Toggle1"
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "XSE"
      BackColor       =   12648447
      ForeColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "KBasicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatScrollBar1_Change()

End Sub

Private Sub FlatScrollBar1_Scroll()

End Sub

Private Sub HScroll3_Change()

End Sub

Private Sub HScroll3_Scroll()

End Sub
