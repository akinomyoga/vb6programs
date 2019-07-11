VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form KBasicForm 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Button"
      TabPicture(0)   =   "KBasicForm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ToggleButton4"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(3)=   "ToggleButton3"
      Tab(0).Control(4)=   "ColorButton1"
      Tab(0).Control(5)=   "ToggleButton2"
      Tab(0).Control(6)=   "ToggleButton1"
      Tab(0).Control(7)=   "ColorButton2"
      Tab(0).Control(8)=   "ColorButton3"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Spin"
      TabPicture(1)   =   "KBasicForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "UpDown1"
      Tab(1).Control(1)=   "UpDown2"
      Tab(1).Control(2)=   "UpDown3"
      Tab(1).Control(3)=   "UpDown4"
      Tab(1).Control(4)=   "SpinButton5"
      Tab(1).Control(5)=   "SpinButton2"
      Tab(1).Control(6)=   "SpinButton4"
      Tab(1).Control(7)=   "SpinButton1"
      Tab(1).Control(8)=   "SpinButton3"
      Tab(1).Control(9)=   "SpinButton6"
      Tab(1).Control(10)=   "UpDown5"
      Tab(1).Control(11)=   "UpDown6"
      Tab(1).Control(12)=   "Label1"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Scroll"
      TabPicture(2)   =   "KBasicForm.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "ScrollBar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FlatScrollBar2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FlatScrollBar1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ScrollBar1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ScrollBar2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "HScroll1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "HScroll3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "HScroll2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin KBasic.ToggleButton ToggleButton4 
         Height          =   375
         Left            =   -71640
         TabIndex        =   21
         ToolTipText     =   "Toggle1"
         Top             =   960
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Command1"
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   120
         Max             =   10
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   120
         Max             =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Max             =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   255
         Left            =   1680
         Top             =   720
         Width           =   1335
         _extentx        =   2355
         _extenty        =   450
         backcolor       =   12632319
      End
      Begin KBasic.ScrollBar ScrollBar1 
         Height          =   135
         Left            =   1680
         Top             =   480
         Width           =   1335
         _extentx        =   2355
         _extenty        =   238
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin KBasic.ScrollBar ScrollBar3 
         Height          =   375
         Left            =   1680
         Top             =   1080
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         max             =   5
         backcolor       =   12648384
         forecolor       =   32768
         enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   -74520
         TabIndex        =   7
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   375
         Left            =   -74520
         TabIndex        =   9
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin KBasic.SpinButton SpinButton5 
         Height          =   375
         Left            =   -73200
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Enabled         =   0   'False
      End
      Begin KBasic.SpinButton SpinButton2 
         Height          =   255
         Left            =   -72840
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Orientation     =   1
      End
      Begin KBasic.SpinButton SpinButton4 
         Height          =   375
         Left            =   -72840
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Orientation     =   1
         BackColor       =   12648384
         ForeColor       =   32768
      End
      Begin KBasic.SpinButton SpinButton1 
         Height          =   255
         Left            =   -73200
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin KBasic.SpinButton SpinButton3 
         Height          =   375
         Left            =   -73200
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   12632319
         ForeColor       =   128
      End
      Begin KBasic.SpinButton SpinButton6 
         Height          =   375
         Left            =   -72840
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Orientation     =   1
         BackColor       =   12648384
         ForeColor       =   32768
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown6 
         Height          =   375
         Left            =   -74520
         TabIndex        =   11
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         Orientation     =   1
         Enabled         =   0   'False
      End
      Begin KBasic.ToggleButton ToggleButton3 
         Height          =   375
         Left            =   -72240
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
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
      Begin KBasic.ColorButton ColorButton1 
         Height          =   375
         Left            =   -73560
         TabIndex        =   15
         ToolTipText     =   "Color1"
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   -72240
         TabIndex        =   16
         ToolTipText     =   "Toggle2"
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   -72240
         TabIndex        =   17
         ToolTipText     =   "Toggle1"
         Top             =   960
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
      Begin KBasic.ColorButton ColorButton2 
         Height          =   375
         Left            =   -73560
         TabIndex        =   19
         ToolTipText     =   "Color1"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   -73560
         TabIndex        =   20
         ToolTipText     =   "Color1"
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   375
         Left            =   -72360
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
   End
End
Attribute VB_Name = "KBasicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SpinButton4_Change()
    Label1.Caption = SpinButton4.Value
End Sub

