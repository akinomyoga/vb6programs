VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#5.0#0"; "KBasic.ocx"
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
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Button"
      TabPicture(0)   =   "KBasicForm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "ColorButton4(7)"
      Tab(0).Control(2)=   "ColorButton3"
      Tab(0).Control(3)=   "ColorButton2"
      Tab(0).Control(4)=   "ToggleButton1"
      Tab(0).Control(5)=   "ToggleButton2"
      Tab(0).Control(6)=   "ColorButton1"
      Tab(0).Control(7)=   "ToggleButton3"
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(10)=   "ToggleButton4"
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(12)=   "ColorButton4(6)"
      Tab(0).Control(13)=   "ColorButton4(5)"
      Tab(0).Control(14)=   "ColorButton4(4)"
      Tab(0).Control(15)=   "ColorButton4(3)"
      Tab(0).Control(16)=   "ColorButton4(2)"
      Tab(0).Control(17)=   "ColorButton4(1)"
      Tab(0).Control(18)=   "ColorButton4(0)"
      Tab(0).Control(19)=   "TemplateControl1"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Spin"
      TabPicture(1)   =   "KBasicForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "UpDown6"
      Tab(1).Control(2)=   "UpDown5"
      Tab(1).Control(3)=   "SpinButton6"
      Tab(1).Control(4)=   "SpinButton3"
      Tab(1).Control(5)=   "SpinButton1"
      Tab(1).Control(6)=   "SpinButton4"
      Tab(1).Control(7)=   "SpinButton2"
      Tab(1).Control(8)=   "SpinButton5"
      Tab(1).Control(9)=   "UpDown4"
      Tab(1).Control(10)=   "UpDown3"
      Tab(1).Control(11)=   "UpDown2"
      Tab(1).Control(12)=   "UpDown1"
      Tab(1).Control(13)=   "ColorButton4(16)"
      Tab(1).Control(14)=   "ColorButton4(15)"
      Tab(1).Control(15)=   "ColorButton4(14)"
      Tab(1).Control(16)=   "ColorButton4(13)"
      Tab(1).Control(17)=   "ColorButton4(12)"
      Tab(1).Control(18)=   "ColorButton4(11)"
      Tab(1).Control(19)=   "ColorButton4(10)"
      Tab(1).Control(20)=   "ColorButton4(9)"
      Tab(1).Control(21)=   "ColorButton4(8)"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Scroll"
      TabPicture(2)   =   "KBasicForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ScrollBar2(4)"
      Tab(2).Control(1)=   "ScrollBar2(3)"
      Tab(2).Control(2)=   "ScrollBar2(2)"
      Tab(2).Control(3)=   "ScrollBar2(1)"
      Tab(2).Control(4)=   "ScrollBar3"
      Tab(2).Control(5)=   "FlatScrollBar2"
      Tab(2).Control(6)=   "FlatScrollBar1"
      Tab(2).Control(7)=   "ScrollBar1"
      Tab(2).Control(8)=   "ScrollBar2(0)"
      Tab(2).Control(9)=   "HScroll1"
      Tab(2).Control(10)=   "HScroll3"
      Tab(2).Control(11)=   "HScroll2"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "KBasicForm.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "ColorButton5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "ToggleButton5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "SpinButton7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "ScrollBar4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TemplateControl2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin KBasic.TemplateControl TemplateControl2 
         Height          =   495
         Left            =   2040
         TabIndex        =   44
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin KBasic.ScrollBar ScrollBar4 
         Height          =   495
         Left            =   1560
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin KBasic.SpinButton SpinButton7 
         Height          =   495
         Left            =   1080
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin KBasic.ToggleButton ToggleButton5 
         Height          =   495
         Left            =   600
         TabIndex        =   43
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "TB"
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   8
         Left            =   -74880
         TabIndex        =   33
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   9
         Left            =   -74520
         TabIndex        =   34
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   10
         Left            =   -74160
         TabIndex        =   35
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   11
         Left            =   -74880
         TabIndex        =   36
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   12
         Left            =   -74520
         TabIndex        =   37
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   13
         Left            =   -74160
         TabIndex        =   38
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   14
         Left            =   -74880
         TabIndex        =   39
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   15
         Left            =   -74520
         TabIndex        =   40
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   16
         Left            =   -74160
         TabIndex        =   41
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.TemplateControl TemplateControl1 
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   23
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   1
         Left            =   -74160
         TabIndex        =   24
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   1
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   2
         Left            =   -73680
         TabIndex        =   25
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   2
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   3
         Left            =   -73200
         TabIndex        =   26
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   3
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   4
         Left            =   -72720
         TabIndex        =   27
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   5
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   5
         Left            =   -72240
         TabIndex        =   28
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   6
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   6
         Left            =   -71760
         TabIndex        =   29
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   7
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Command1"
         Height          =   375
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin KBasic.ToggleButton ToggleButton4 
         Height          =   375
         Left            =   -71640
         TabIndex        =   21
         ToolTipText     =   "Toggle1"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BackColor       =   12648447
         ForeColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "XSE"
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
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   -74880
         Max             =   10
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   -74880
         Max             =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         Max             =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   255
         Index           =   0
         Left            =   -73320
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   12632319
      End
      Begin KBasic.ScrollBar ScrollBar1 
         Height          =   135
         Left            =   -73320
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Max             =   10
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   2
         Arrows          =   65536
         Max             =   10
         Orientation     =   1638401
      End
      Begin KBasic.ScrollBar ScrollBar3 
         Height          =   375
         Left            =   -73320
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Enabled         =   0   'False
         BackColor       =   12648384
         ForeColor       =   32768
         Max             =   5
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
         BackColor       =   12648384
         ForeColor       =   32768
         Orientation     =   1
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
         Enabled         =   0   'False
         BackColor       =   12648384
         ForeColor       =   32768
         Orientation     =   1
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
         Value           =   -1  'True
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
         Caption         =   "XYZ"
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Color2"
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
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   255
         Index           =   1
         Left            =   -73320
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   12632319
         Appearance      =   1
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   255
         Index           =   2
         Left            =   -73320
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   12632319
         Delay           =   100
         Appearance      =   2
      End
      Begin KBasic.ColorButton ColorButton4 
         Height          =   375
         Index           =   7
         Left            =   -71280
         TabIndex        =   31
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "A"
         Appearance      =   8
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   1695
         Index           =   3
         Left            =   -71880
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2990
         BackColor       =   16761024
         ButtonSize      =   15
         Appearance      =   2
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   1695
         Index           =   4
         Left            =   -71400
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2990
         BackColor       =   16761024
         Appearance      =   2
      End
      Begin KBasic.ColorButton ColorButton5 
         Height          =   495
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Color1"
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CB"
      End
      Begin VB.Label Label2 
         Caption         =   "Button Appearances"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
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

