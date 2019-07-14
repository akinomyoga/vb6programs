VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#5.0#0"; "KBasic.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form KBasicForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Test KBasic Controls"
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
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Button"
      TabPicture(0)   =   "KBasicForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "KButton4(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "KButton4(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "KButton4(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "KButton3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "KButton2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "KButton1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "KButton4(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "KButton4(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "KButton4(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "KButton4(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "KButton4(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "KButton4(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "KButton4(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Spin"
      TabPicture(1)   =   "KBasicForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "SpinButton8(9)"
      Tab(1).Control(3)=   "SpinButton8(8)"
      Tab(1).Control(4)=   "SpinButton8(7)"
      Tab(1).Control(5)=   "SpinButton8(6)"
      Tab(1).Control(6)=   "SpinButton8(5)"
      Tab(1).Control(7)=   "SpinButton8(4)"
      Tab(1).Control(8)=   "SpinButton8(3)"
      Tab(1).Control(9)=   "SpinButton8(2)"
      Tab(1).Control(10)=   "SpinButton8(1)"
      Tab(1).Control(11)=   "UpDown6"
      Tab(1).Control(12)=   "UpDown5"
      Tab(1).Control(13)=   "SpinButton6"
      Tab(1).Control(14)=   "SpinButton3(0)"
      Tab(1).Control(15)=   "SpinButton1"
      Tab(1).Control(16)=   "SpinButton4"
      Tab(1).Control(17)=   "SpinButton2"
      Tab(1).Control(18)=   "SpinButton5"
      Tab(1).Control(19)=   "UpDown4"
      Tab(1).Control(20)=   "UpDown3"
      Tab(1).Control(21)=   "UpDown2"
      Tab(1).Control(22)=   "UpDown1"
      Tab(1).Control(23)=   "SpinButton8(0)"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Scroll"
      TabPicture(2)   =   "KBasicForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ScrollBar6(0)"
      Tab(2).Control(1)=   "VScroll3"
      Tab(2).Control(2)=   "VScroll2"
      Tab(2).Control(3)=   "VScroll1"
      Tab(2).Control(4)=   "ScrollBar2(0)"
      Tab(2).Control(5)=   "ScrollBar1"
      Tab(2).Control(6)=   "FlatScrollBar1"
      Tab(2).Control(7)=   "ScrollBar3"
      Tab(2).Control(8)=   "ScrollBar2(1)"
      Tab(2).Control(9)=   "ScrollBar2(3)"
      Tab(2).Control(10)=   "ScrollBar2(4)"
      Tab(2).Control(11)=   "ScrollBar6(1)"
      Tab(2).Control(12)=   "ScrollBar6(2)"
      Tab(2).Control(13)=   "ScrollBar6(3)"
      Tab(2).Control(14)=   "ScrollBar6(5)"
      Tab(2).Control(15)=   "ScrollBar6(6)"
      Tab(2).Control(16)=   "ScrollBar6(4)"
      Tab(2).Control(17)=   "ScrollBar6(7)"
      Tab(2).Control(18)=   "ScrollBar6(8)"
      Tab(2).Control(19)=   "ScrollBar6(9)"
      Tab(2).Control(20)=   "FlatScrollBar2"
      Tab(2).Control(21)=   "Label4"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Toggle"
      TabPicture(3)   =   "KBasicForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check1(2)"
      Tab(3).Control(1)=   "Check1(1)"
      Tab(3).Control(2)=   "Check1(0)"
      Tab(3).Control(3)=   "ScrollBar4"
      Tab(3).Control(4)=   "SpinButton7"
      Tab(3).Control(5)=   "ToggleButton5"
      Tab(3).Control(6)=   "KButton5"
      Tab(3).Control(7)=   "ToggleButton4"
      Tab(3).Control(8)=   "ToggleButton3"
      Tab(3).Control(9)=   "ToggleButton2"
      Tab(3).Control(10)=   "ToggleButton1"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Progress"
      TabPicture(4)   =   "KBasicForm.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Timer1"
      Tab(4).Control(1)=   "KProgressBar1"
      Tab(4).Control(2)=   "ProgressBar1"
      Tab(4).ControlCount=   3
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   -72840
         Top             =   840
      End
      Begin KBasic.KProgressBar KProgressBar1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Check1"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   0
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   1215
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   0
         Left            =   -74760
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1095
         Left            =   -74400
         Max             =   10
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1095
         Left            =   -74640
         Max             =   10
         TabIndex        =   29
         Top             =   480
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1095
         Left            =   -74880
         Max             =   10
         TabIndex        =   28
         Top             =   480
         Width           =   135
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   0
         Left            =   -74520
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
      End
      Begin KBasic.ScrollBar ScrollBar4 
         Height          =   495
         Left            =   -73440
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin KBasic.SpinButton SpinButton7 
         Height          =   495
         Left            =   -73920
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin KBasic.ToggleButton ToggleButton5 
         Height          =   495
         Left            =   -74400
         TabIndex        =   25
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "TB"
      End
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   16
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   18
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   5
         Left            =   840
         TabIndex        =   20
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   21
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   1095
         Index           =   0
         Left            =   -73080
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1931
         BackColor       =   12632319
      End
      Begin KBasic.ScrollBar ScrollBar1 
         Height          =   1095
         Left            =   -73320
         Top             =   480
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1931
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   1095
         Left            =   -73920
         TabIndex        =   1
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1931
         _Version        =   393216
         Max             =   10
         Orientation     =   1638400
      End
      Begin KBasic.ScrollBar ScrollBar3 
         Height          =   1095
         Left            =   -72720
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1931
         Enabled         =   0   'False
         BackColor       =   12648384
         ForeColor       =   32768
         Max             =   5
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         Index           =   0
         Left            =   -73200
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   12632319
         ForeColor       =   128
         Appearance      =   7
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         Orientation     =   1
         Enabled         =   0   'False
      End
      Begin KBasic.KButton KButton1 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
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
      Begin KBasic.KButton KButton2 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
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
      Begin KBasic.KButton KButton3 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
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
         Height          =   1095
         Index           =   1
         Left            =   -72240
         ToolTipText     =   "Slow ScrollBar"
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1931
         BackColor       =   12632319
         Delay           =   100
         Appearance      =   1
      End
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   23
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
         Height          =   1095
         Index           =   3
         Left            =   -71880
         ToolTipText     =   "button/bar size specified"
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1931
         BackColor       =   16761024
         ButtonSize      =   9
         Appearance      =   3
      End
      Begin KBasic.ScrollBar ScrollBar2 
         Height          =   1095
         Index           =   4
         Left            =   -71400
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1931
         BackColor       =   16761024
         BarSize         =   15
         ButtonSize      =   15
         Appearance      =   3
      End
      Begin KBasic.KButton KButton5 
         Height          =   495
         Left            =   -74880
         TabIndex        =   24
         ToolTipText     =   "Color1"
         Top             =   2520
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   17
         Left            =   2280
         TabIndex        =   26
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
         Appearance      =   9
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   1
         Left            =   -74160
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   1
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   2
         Left            =   -73800
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   2
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   3
         Left            =   -73440
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   3
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   4
         Left            =   -73080
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   4
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   5
         Left            =   -72720
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   5
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   6
         Left            =   -72360
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   6
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   7
         Left            =   -72000
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   7
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   8
         Left            =   -71640
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   8
      End
      Begin KBasic.SpinButton SpinButton8 
         Height          =   375
         Index           =   9
         Left            =   -71280
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   9
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   1
         Left            =   -74400
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   1
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   2
         Left            =   -74040
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   2
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   3
         Left            =   -73680
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   3
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   5
         Left            =   -73320
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   4
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   6
         Left            =   -72960
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   5
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   4
         Left            =   -72600
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   6
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   7
         Left            =   -72240
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   7
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   8
         Left            =   -71880
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   8
      End
      Begin KBasic.ScrollBar ScrollBar6 
         Height          =   975
         Index           =   9
         Left            =   -71520
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         BackColor       =   16761024
         ForeColor       =   16711680
         Appearance      =   9
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
         Height          =   1095
         Left            =   -73680
         TabIndex        =   32
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1931
         _Version        =   393216
         Appearance      =   2
         Max             =   10
         Orientation     =   1638400
      End
      Begin KBasic.ToggleButton ToggleButton4 
         Height          =   375
         Left            =   -72960
         TabIndex        =   33
         ToolTipText     =   "Toggle1"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
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
         Caption         =   "XSE"
      End
      Begin KBasic.ToggleButton ToggleButton3 
         Height          =   375
         Left            =   -73560
         TabIndex        =   34
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
      Begin KBasic.ToggleButton ToggleButton2 
         Height          =   375
         Left            =   -73560
         TabIndex        =   35
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
         Left            =   -73560
         TabIndex        =   36
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
      Begin KBasic.KButton KButton4 
         Height          =   375
         Index           =   8
         Left            =   2280
         TabIndex        =   40
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   16711680
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
         Appearance      =   4
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   10
      End
      Begin VB.Label Label4 
         Caption         =   "ScrollBar Appearances"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "SpinButton Appearances"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Button Appearances"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   -72360
         TabIndex        =   8
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
Private Sub increment_ProgressBar1()
    Dim new_Value As Long
    new_Value = ProgressBar1.Value + 1
    If new_Value > ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Min
    Else
        ProgressBar1.Value = new_Value
    End If
End Sub

Private Sub increment_KProgressBar1()
    If KProgressBar1.Value < KProgressBar1.Max Then
        KProgressBar1.Value = KProgressBar1.Value + 3
    Else
        KProgressBar1.Value = KProgressBar1.Min
    End If
    KProgressBar1.Caption = KProgressBar1.Value & "%"
End Sub

Private Sub SpinButton4_Change()
    Label1.Caption = SpinButton4.Value
End Sub

Private Sub Timer1_Timer()
    increment_ProgressBar1
    increment_KProgressBar1
End Sub
