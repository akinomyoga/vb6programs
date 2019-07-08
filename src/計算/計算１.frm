VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#4.0#0"; "KBasic.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "åvéZã@"
   ClientHeight    =   5010
   ClientLeft      =   4650
   ClientTop       =   4200
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
      Size            =   8.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "åvéZÇP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "åvéZ1ÉtÉ@ÉCÉã(*.åvéZ1)|*.åvéZ1"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Å®|"
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   90
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Å®"
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   89
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ".H."
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   83
      ToolTipText     =   "èdï°ëgçáÇπ"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ".C."
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   82
      ToolTipText     =   "ëgçáÇπ"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harccot"
      Height          =   255
      Index           =   29
      Left            =   2280
      TabIndex        =   80
      ToolTipText     =   "ëoã»ê¸ãtó]ê⁄"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harcsec"
      Height          =   255
      Index           =   28
      Left            =   2280
      TabIndex        =   79
      ToolTipText     =   "ëoã»ê¸ãtê≥äÑ"
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harccosec"
      Height          =   255
      Index           =   27
      Left            =   2280
      TabIndex        =   78
      ToolTipText     =   "ëoã»ê¸ãtó]äÑ"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harctan"
      Height          =   255
      Index           =   26
      Left            =   1560
      TabIndex        =   77
      ToolTipText     =   "ëoã»ê¸ãtê≥ê⁄"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harccos"
      Height          =   255
      Index           =   25
      Left            =   1560
      TabIndex        =   76
      ToolTipText     =   "ëoã»ê¸ãtó]å∑"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "harcsin"
      Height          =   255
      Index           =   24
      Left            =   1560
      TabIndex        =   75
      ToolTipText     =   "ëoã»ê¸ãtê≥å∑"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hcot"
      Height          =   255
      Index           =   23
      Left            =   840
      TabIndex        =   74
      ToolTipText     =   "ëoã»ê¸ó]ê⁄"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hsec"
      Height          =   255
      Index           =   22
      Left            =   840
      TabIndex        =   73
      ToolTipText     =   "ëoã»ê¸ê≥äÑ"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hcosec"
      Height          =   255
      Index           =   21
      Left            =   840
      TabIndex        =   72
      ToolTipText     =   "ëoã»ê¸ó]äÑ"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "htan"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   71
      ToolTipText     =   "ëoã»ê¸ê≥ê⁄"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hcos"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   70
      ToolTipText     =   "ëoã»ê¸ó]å∑"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hsin"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   69
      ToolTipText     =   "ëoã»ê¸ê≥å∑"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arccot"
      Height          =   255
      Index           =   17
      Left            =   2280
      TabIndex        =   67
      ToolTipText     =   "ãtê≥ê⁄"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arcsec"
      Height          =   255
      Index           =   16
      Left            =   2280
      TabIndex        =   66
      ToolTipText     =   "ãtê≥äÑ"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arccosec"
      Height          =   255
      Index           =   15
      Left            =   2280
      TabIndex        =   65
      ToolTipText     =   "ãtó]äÑ"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arctan"
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   64
      ToolTipText     =   "ãtê≥ê⁄"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arccos"
      Height          =   255
      Index           =   13
      Left            =   1560
      TabIndex        =   63
      ToolTipText     =   "ãtó]å∑"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "arcsin"
      Height          =   255
      Index           =   12
      Left            =   1560
      TabIndex        =   62
      ToolTipText     =   "ãtê≥å∑"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cot"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   61
      ToolTipText     =   "ó]ê⁄"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sec"
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   60
      ToolTipText     =   "ê≥äÑ"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "tan"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   59
      ToolTipText     =   "ê≥ê⁄"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cos"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   58
      ToolTipText     =   "ó]å∑"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "| |"
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   55
      ToolTipText     =   "ê‚ëŒíl"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Å©"
      Height          =   375
      Index           =   9
      Left            =   3240
      TabIndex        =   54
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Åc"
      Height          =   375
      Index           =   9
      Left            =   3000
      TabIndex        =   53
      ToolTipText     =   "èËó]"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1/"
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   52
      ToolTipText     =   "ãtêî"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Åì"
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   51
      ToolTipText     =   "ïSï™ó¶"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CM"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   50
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CM"
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   49
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CM"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   48
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CM"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   47
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RM"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   46
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RM"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   45
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RM"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   44
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RM"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   43
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M-"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   42
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M-"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   41
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M-"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   40
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M-"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   39
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M+"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   38
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M+"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   37
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M+"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M+"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   35
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "M"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   30
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "M"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   29
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "M"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   28
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "M"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   27
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+-"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   26
      ToolTipText     =   "ïÑçÜ"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   25
      ToolTipText     =   "Clear"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÅO"
      Height          =   375
      Index           =   7
      Left            =   3000
      TabIndex        =   23
      ToolTipText     =   "ó›èÊ"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "="
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   22
      ToolTipText     =   "åãâ "
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "!"
      Height          =   375
      Index           =   5
      Left            =   2040
      TabIndex        =   21
      ToolTipText     =   "äKèÊ"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉŒ"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   20
      ToolTipText     =   "â~é¸ó¶"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "00"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Å„"
      Height          =   375
      Index           =   4
      Left            =   3000
      TabIndex        =   18
      ToolTipText     =   "ó›èÊç™"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÅÄ"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   17
      ToolTipText     =   "èúéZ"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Å~"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   16
      ToolTipText     =   "èÊéZ"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Å|"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   15
      ToolTipText     =   "å∏éZ"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Å{"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "â¡éZ"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "AC"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "All Clear"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "è¨êîì_"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cosec"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   68
      ToolTipText     =   "ó]äÑ"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sin"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   57
      ToolTipText     =   "ê≥å∑"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Åõ"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   85
      ToolTipText     =   "â~èáóÒ"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ";;;;;;;;''''''''"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   84
      ToolTipText     =   "èdï°èáóÒ"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ";;;;;,,,::::  "
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   86
      ToolTipText     =   "ìØÇ∂ï®Çä‹ÇﬁèáóÒ"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   ".P."
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   81
      ToolTipText     =   "èáóÒ"
      Top             =   3480
      Width           =   375
   End
   Begin KBasic.ToggleButton ToggleButton1 
      Height          =   315
      Left            =   360
      TabIndex        =   56
      ToolTipText     =   "äpìxÇÃíPà "
      Top             =   2880
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "ÉâÉWÉAÉì"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      Caption         =   "éOäpä÷êî"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   92
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      Caption         =   "èÍçáÇÃêî"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3960
      TabIndex        =   91
      Top             =   2880
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      Height          =   1335
      Left            =   3480
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   88
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   4080
      TabIndex        =   87
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   1935
      Left            =   60
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1935
      Left            =   3960
      Top             =   900
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   4080
      TabIndex        =   34
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   4080
      TabIndex        =   33
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   4080
      TabIndex        =   32
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   4080
      TabIndex        =   31
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   600
      TabIndex        =   13
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      TabIndex        =   10
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu mnfile 
      Caption         =   "ÉtÉ@ÉCÉã"
      Begin VB.Menu mnsave 
         Caption         =   "ç°ÇÃèÛë‘Çï€ë∂..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnopen 
         Caption         =   "ï€ë∂ÇµÇΩï®ÇäJÇ≠..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnend 
         Caption         =   "èIóπ"
      End
   End
   Begin VB.Menu mnhoka 
      Caption         =   "ëºÇÃåvéZ"
      Begin VB.Menu mntani 
         Caption         =   "íPà ÇÃïœä∑..."
      End
      Begin VB.Menu mnsankaku 
         Caption         =   "éOäpå`í≤Ç◊..."
      End
   End
   Begin VB.Menu mnhyou 
      Caption         =   "ï\"
      Begin VB.Menu root 
         Caption         =   "ïΩï˚ç™ï\"
      End
      Begin VB.Menu kuku 
         Caption         =   "ã„ã„ï\"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "ÉwÉãÉv"
      Begin VB.Menu mnh1 
         Caption         =   "É{É^Éì"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a1 As Double, a2 As Double, b1 As Integer, b2 As Integer, c1 As Integer, c2 As Integer
Dim d1(1) As Integer
Public path1

Private Sub Command1_Click(Index As Integer)
    If Label1.Caption = "0" Then Label1.Caption = ""
    Label1.Caption = Label1.Caption & Index
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim X As Double
    On Error GoTo err
    Select Case Index
    Case 0
        Label1.Caption = Int(Label1.Caption) & "."
    Case 1
        If Label1.Caption <> 0 Then Label1.Caption = Label1.Caption & "00"
    Case 2
        Label1.Caption = "3.14159265358979323846"
    Case 3
        Label1.Caption = -Label1.Caption
    Case 4
        Label1.Caption = 1 / Label1.Caption
    Case 5
        Label1.Caption = Abs(Label1.Caption)
    Case 6 To 11
        If ToggleButton1.Value = True Then Label1.Caption = Label1.Caption / 180 * 3.14159265358979
        Select Case Index
        Case 6
            Label1.Caption = Sin(Label1.Caption)
        Case 7
            Label1.Caption = Cos(Label1.Caption)
        Case 8
            Label1.Caption = Tan(Label1.Caption)
        Case 9
            Label1.Caption = 1 / Sin(Label1.Caption)
        Case 10
            Label1.Caption = 1 / Cos(Label1.Caption)
        Case 11
            Label1.Caption = 1 / Tan(Label1.Caption)
        End Select
    Case 12 To 17
        X = Label1.Caption
        Select Case Index
        Case 12
            Label1.Caption = Atn(X / Sqr(-X * X + 1))
        Case 13
            Label1.Caption = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
        Case 14
            Label1.Caption = Atn(X)
        Case 16
            Label1.Caption = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
        Case 15
            Label1.Caption = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
        Case 17
            Label1.Caption = Atn(X) + 2 * Atn(1)
        End Select
        If ToggleButton1.Value = True Then Label1.Caption = Label1.Caption * 180 / 3.14159265358979
    Case 18 To 23
        If ToggleButton1.Value = True Then Label1.Caption = Label1.Caption / 180 * 3.14159265358979
        X = Label.Caption
        Select Case Index
        Case 18
            Label1.Caption = (Exp(X) - Exp(-X)) / 2
        Case 19
            Label1.Caption = (Exp(X) + Exp(-X)) / 2
        Case 20
            Label1.Caption = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
        Case 22
            Label1.Caption = 2 / (Exp(X) + Exp(-X))
        Case 21
            Label1.Caption = 2 / (Exp(X) - Exp(-X))
        Case 23
            Label1.Caption = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
        End Select
    Case 24 To 29
        X = Label.Caption
        Select Case Index
        Case 24
            Label1.Caption = Log(X + Sqr(X * X + 1))
        Case 25
            Label1.Caption = Log(X + Sqr(X * X - 1))
        Case 26
            Label1.Caption = Log((1 + X) / (1 - X)) / 2
        Case 28
            Label1.Caption = Log((Sqr(-X * X + 1) + 1) / X)
        Case 27
            Label1.Caption = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
        Case 29
            Label1.Caption = Log((X + 1) / (X - 1)) / 2
        End Select
        If ToggleButton1.Value = True Then Label1.Caption = Label1.Caption * 180 / 3.14159265358979
    End Select
    Exit Sub
err:
    If err.Number = 11 Then
        MsgBox "0Ç≈åvéZÇ≈Ç´Ç‹ÇπÇÒÅI"
    ElseIf err.Number = 5 Then
        MsgBox "óLÇËìæÇÈêîílÇì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅI" & Chr(13) & "arcsin[cos]:0-1" & Chr(13) & "arccosec[sec]:1-"
    Else
        MsgBox "âΩÇ©ÇÃÉGÉâÅ[Ç™ãNÇ´ÇƒåvéZÇ≈Ç´Ç‹ÇπÇÒÅBÉGÉâÅ[ÉiÉìÉoÅ[" & err.Number
        Call Command3_Click
    End If
End Sub

Private Sub Command3_Click()
    Label1.Caption = 0
    a1 = 0: a2 = 0
    Label2.Caption = ""
    b1 = 0: b2 = 0
    Label3.Caption = ""
    For a = 0 To 3
        Label4(a).Caption = "0"
    Next a
    Label5.Caption = ""
    Label6.Caption = ""
End Sub

Private Sub Command4_Click(Index As Integer)
    On Error GoTo err
    b3 = b2
    If b2 = 1 Then a2 = Label1.Caption
    If b2 = 0 And Index <> 6 Then GoTo skip1
    If Index = 5 Then GoTo skip1
    b2 = 0
    Select Case b1
    Case 1
        Label1.Caption = a1 + a2
    Case 2
        Label1.Caption = a1 - a2
    Case 3
        Label1.Caption = a1 * a2
    Case 4
        Label1.Caption = a1 / a2
    Case 5
        Label1.Caption = a1 ^ a2
    Case 6
        Label1.Caption = a1 ^ (1 / a2)
    Case 7
        Label1.Caption = a1 Mod a2
    End Select
skip1:
    Select Case Index
    Case 0
        Label3.Caption = "Å{"
        Label2.Caption = Label1.Caption & "Å{ ? =[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 1
        b2 = 1
    Case 1
        Label3.Caption = "Å|"
        Label2.Caption = Label1.Caption & "Å| ? =[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 2
        b2 = 1
    Case 2
        Label3.Caption = "Å~"
        Label2.Caption = Label1.Caption & "Å~ ? =[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 3
        b2 = 1
    Case 3
        Label3.Caption = "ÅÄ"
        Label2.Caption = Label1.Caption & "ÅÄ ? =[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 4
        b2 = 1
    Case 7
        Label3.Caption = "ÅO"
        Label2.Caption = Label1.Caption & "ÅO ? =[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 5
        b2 = 1
    Case 4
        Label3.Caption = "Å„"
        Label2.Caption = "? Å„" & Label1.Caption & "=[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 6
        b2 = 1
    Case 9
        Label3.Caption = "Åc"
        Label2.Caption = Label1.Caption & "ÅÄ ?=Å†Åc[ ]"
        a1 = Label1.Caption
        Label1.Caption = 0
        b1 = 7
        b2 = 1
    Case 6
        Label2.Caption = a2
        a1 = Label1.Caption
    Case 5
        Label1.Caption = exclam(Int(Label1.Caption))
    Case 8
        If b3 = 1 Then Label1.Caption = 0.01 * Label1.Caption
    End Select
    Exit Sub
err:
    If err.Number = 6 Then
        MsgBox "êîÇ™ëÂÇ´Ç∑Ç¨ÇƒèoóàÇ‹ÇπÇÒÅB"
    Else
        MsgBox "ïsñæÇ»ÉGÉâÅ[Ç≈èoóàÇ‹ÇπÇÒÅB"
    End If
End Sub

Private Sub Command5_Click(Index As Integer)
    Select Case Index
    Case 0
        Label1.Caption = 0
    Case Is < 5
        Label1.Caption = Label4(Index - 1).Caption
    Case Is < 9
        Label4(Index - 5).Caption = 0
    Case 9
        Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1)
        If Label1.Caption = "" Then Label1.Caption = 0
    End Select
End Sub

Public Function exclam(a As Integer)
    c = 1
    If a > 0 Then
        For b = 1 To a
            c = c * b
        Next b
    ElseIf a < 0 Then
        For b = -1 To a Step -1
            c = c * b
        Next b
    End If
    exclam = c
End Function

Private Sub Command6_Click(Index As Integer)
    Label4(Index).Caption = Label1.Caption
End Sub

Private Sub Command7_Click(Index As Integer)
    If Index < 4 Then
        Label4(Index).Caption = Label4(Index).Caption - -Label1.Caption
    Else
        Label4(Index - 4).Caption = Label4(Index - 4).Caption - Label1.Caption
    End If
End Sub

Private Sub Command8_Click(Index As Integer)
    On Error GoTo err
    Select Case Index
    Case 0 To 5
        Label6.Caption = Command8(Index).Caption
        Label6.ToolTipText = Command8(Index).ToolTipText
        If Index = 2 Then
            Label6.Caption = ""
            Label6.ToolTipText = ""
            Label1.Caption = exclam(Label1.Caption - 1)
            d1(1) = 1
        Else
            c1 = Index
            d1(0) = Label1.Caption
            Label5.Caption = d1(0)
            c2 = 0
            Label1.Caption = 0
        End If
    Case 6
        c2 = 1
        d1(c2) = Label1.Caption
        Select Case c1
        Case 3
            Label5.Caption = Label5.Caption & "," & Int(Label1.Caption)
            d1(c2) = d1(c2) * exclam(Label1.Caption)
            d1(0) = d1(0) - -Label1.Caption
            Label1.Caption = 0
            Exit Sub
        Case 0
            Label1.Caption = Int(exclam(d1(0)) / exclam(d1(0) - Label1.Caption))
        Case 1
            Label1.Caption = d1(0) ^ Int(Label1.Caption)
        Case 4
            Label1.Caption = Int(exclam(d1(0)) / exclam(d1(0) - Label1.Caption) / exclam(Label1.Caption))
        Case 5
            Label1.Caption = exclam(d1(0) + Label1.Caption - 1) / exclam(d1(0) - 1) / exclam(Label1.Caption)
        End Select
        If Label1.Caption = 0 Then MsgBox "êîílÇämÇ©ÇﬂÇƒÇ≠ÇæÇ≥Ç¢ÅI"
        c2 = 0
        Label6.Caption = ""
        Label6.ToolTipText = ""
        Label5.Caption = ""
        d1(0) = 0
        d1(1) = 0
    Case 7
        If c1 = 3 Then
            c2 = 0
            Label6.Caption = ""
            Label6.ToolTipText = ""
            Label5.Caption = ""
            Label1.Caption = Int(exclam(d1(0)) / d1(1))
            If Label1.Caption = 0 Then MsgBox "êîílÇämÇ©ÇﬂÇƒÇ≠ÇæÇ≥Ç¢ÅI"
            c2 = 0
            Label6.Caption = ""
            Label6.ToolTipText = ""
            Label5.Caption = ""
            d1(0) = 0
            d1(1) = 0
        End If
    End Select
    Exit Sub
err:
    MsgBox err.Number
End Sub

Private Sub Form_Load()
    path1 = CurDir
    CommonDialog1.InitDir = path1
End Sub

Private Sub kuku_Click()
    Open path1 & "\ã„ã„ï\.txt" For Output As 1
    Print #1, "     *ã„ã„ï\*"
    Print #1, "*| 1| 2| 3| 4| 5| 6| 7| 8| 9"
    For a = 1 To 9
        c = a
        For b = 1 To 9
            c = c & "|" & Format(a * b, "00")
        Next b
        Print #1, c
    Next a
    Print #1, "   ****åvéZÇP made by ë∫ê£å˜àÍ****"
    Close #1
    MsgBox path1 & "\ã„ã„ï\.txtÇ…ï€ë∂Ç≥ÇÍÇ‹ÇµÇΩÅI"
End Sub

Private Sub mnend_Click()
    End
End Sub

Private Sub mnh1_Click()
    MsgBox "[1-9][0][00][.]:êîéöÇì¸óÕÇ∑ÇÈÉ{É^Éì" & Chr(13) & "[ÉŒ]:â~é¸ó¶Çì¸óÕ" _
    & Chr(13) & "[| |]:ï\é¶Ç≥ÇÍÇƒÇ¢ÇÈêîílÇê‚ëŒílÇ…Ç∑ÇÈ" & Chr(13) & "[1/]:ï\é¶Ç≥ÇÍÇƒÇ¢ÇÈêîílÇãtêîÇ…Ç∑ÇÈ" _
    & Chr(13) & "[+-]:ï\é¶Ç≥ÇÍÇƒÇ¢ÇÈêîílÇÃïÑçÜÇïœÇ¶ÇÈ" & Chr(13) & "[!]:ï\é¶Ç≥ÇÍÇƒÇ¢ÇÈêîílÅiéléÃå‹ì¸ÇµÇΩílÅjÇäKèÊÇ∑ÇÈ" _
    & Chr(13) & "[+][-][Å~][ÅÄ]:â¡å∏èÊèúÇÃåvéZÇÇ∑ÇÈ" & Chr(13) & "[Å„]:ó›èÊç™Ç∑ÇÈÅiâΩèÊç™Ç©ì¸óÕÇµÇƒâ∫Ç≥Ç¢Åj" _
    & Chr(13) & "[^]:ó›èÊÇ…Ç∑ÇÈ" & Chr(13) & "[Åc]:äÑÇËéZÇÃó]ÇËÇãÅÇﬂÇÈ" _
    & Chr(13) & "[Åì]:ïSï™ó¶Ç≈ï\Ç∑(1/100)" & Chr(13) & "[=]:ìöÇ¶ÇèoÇ∑" _
    & Chr(13) & "[Å©]:êîéöÇàÍÇ¬è¡Ç∑" & Chr(13) & "[C]:ï\é¶Ç≥ÇÍÇƒÇ¢ÇÈêîéöÇè¡ÇµÅAÇOÇ…Ç∑ÇÈ" _
    & Chr(13) & "[AC]:ëSÇƒÇÃílÇénÇﬂÇ…ñﬂÇ∑"
End Sub

Private Sub mnopen_Click()
    On Error GoTo errhand
    CommonDialog1.Filter = "åvéZ1ÉtÉ@ÉCÉã(*.åvéZÇP)|*.åvéZ1"
    CommonDialog1.ShowOpen
    a = CommonDialog1.FileName
    If Right(a, 4) <> ".åvéZ1" Then a = a & ".åvéZ1"
    Dim b(4) As Double
    Dim c(3)
    Open a For Input As 1
    Line Input #1, d
    Line Input #1, e
    Line Input #1, f
    Line Input #1, g
    Line Input #1, h
    Line Input #1, i
    Line Input #1, j
    Line Input #1, k
    Line Input #1, l
    Line Input #1, M
    Line Input #1, n
    Line Input #1, o
    Line Input #1, p
    Line Input #1, q
    Line Input #1, r
    Line Input #1, s
    Line Input #1, t
    Close #1
    b(0) = d: Label1.Caption = b(0)
    b(1) = g: Label4(0).Caption = b(1)
    b(2) = h: Label4(1).Caption = b(2)
    b(3) = i: Label4(2).Caption = b(3)
    b(4) = j: Label4(3).Caption = b(4)
    Label2.Caption = e: Label3.Caption = f
    Label5.Caption = k: Label6.Caption = l
    a1 = M: a2 = n: b1 = o: b2 = p
    c1 = q: c2 = r: d1(0) = s: d1(1) = t
    Exit Sub
errhand:
    MsgBox "ê≥ÇµÇ¢ÉtÉ@ÉCÉãÇëIÇÒÇ≈Ç≠ÇæÇ≥Ç¢ÅI"
End Sub


Private Sub mnsankaku_Click()
    Form3.Show
End Sub

Private Sub mnsave_Click()
    On Error GoTo err
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "åvéZ1ÉtÉ@ÉCÉã(*.åvéZÇP)|*.åvéZ1"
    CommonDialog1.ShowSave
    a = CommonDialog1.FileName
    If Right(a, 4) <> ".åvéZ1" Then a = a & ".åvéZ1"
    Open a For Output As 1
    Print #1, Label1.Caption
    Print #1, Label2.Caption
    Print #1, Label3.Caption
    Print #1, Label4(0).Caption
    Print #1, Label4(1).Caption
    Print #1, Label4(2).Caption
    Print #1, Label4(3).Caption
    Print #1, Label5.Caption
    Print #1, Label6.Caption
    Print #1, a1
    Print #1, a2
    Print #1, b1
    Print #1, b2
    Print #1, c1
    Print #1, c2
    Print #1, d1(0)
    Print #1, d1(1)
    Close #1
err:
    If err.Number = 72755 Then
        MsgBox "ÉÜÅ[ÉUÅ[Ç…ÇÊÇËÉLÉÉÉìÉZÉãÇ≥ÇÍÇ‹ÇµÇΩ!", 64, "Canceled"
    Else
        MsgBox "error", 16, "erred"
    End If
End Sub

Private Sub mntani_Click()
    Form2.Show
End Sub

Private Sub root_Click()
    Open path1 & "\ïΩï˚ç™ï\.txt" For Output As 1
    Print #1, "     *ïΩï˚ç™ï\* è¨êîì_à»â∫ëÊéOà Ç‹Ç≈"
    Print #1, "   |   0|   1|   2|   3|   4|   5|   6|   7|   8|   9"
    For a = 1 To 9.9 Step 0.1
        a1 = Int(a * 10) / 10
        c = a1
        If Len(c) = 1 Then c = c & ".0"
        For b = a1 To a1 + 0.09 Step 0.01
            c = c & "|" & Int(Sqr(b) * 1000)
        Next b
        Print #1, c
    Next a
    For a = 100 To 990 Step 10
        c = " " & a / 10
        a1 = a
        For b = a1 To a1 + 9
            c = c & "|" & Int(Sqr(b / 10) * 1000)
        Next b
        Print #1, c
    Next a
    Print #1, "   ****åvéZÇP made by ë∫ê£å˜àÍ****"
    Close #1
    MsgBox path1 & "\ïΩï˚ç™ï\.txtÇ…ï€ë∂Ç≥ÇÍÇ‹ÇµÇΩÅI"
End Sub

Private Sub ToggleButton1_Click()
    If ToggleButton1.Value = True Then
        ToggleButton1.Caption = "Åã"
    Else
        ToggleButton1.Caption = "ÉâÉWÉAÉì"
    End If
End Sub

