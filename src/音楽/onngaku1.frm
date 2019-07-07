VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "âπäyÇP"
   ClientHeight    =   8910
   ClientLeft      =   2400
   ClientTop       =   1725
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  '’∞ªﬁ∞íËã`
   ScaleHeight     =   8910
   ScaleWidth      =   10170
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   7080
      TabIndex        =   80
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Picture         =   "onngaku1.frx":0000
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "ìríÜÇ©ÇÁçƒê∂"
      Top             =   3480
      Width           =   375
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   5
      Left            =   120
      Max             =   9999
      TabIndex        =   78
      Top             =   8280
      Width           =   7455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Ç»Çµ
      Height          =   3855
      Left            =   120
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      ScaleHeight     =   3855
      ScaleWidth      =   7455
      TabIndex        =   77
      Top             =   4320
      Width           =   7455
   End
   Begin VB.Timer Timer2 
      Left            =   7200
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'â∫ëµÇ¶
      Height          =   375
      Left            =   0
      TabIndex        =   76
      Top             =   8535
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
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
      Left            =   2280
      Picture         =   "onngaku1.frx":014A
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "í‚é~"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Picture         =   "onngaku1.frx":0294
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "àÍéûí‚é~"
      Top             =   3480
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   7110
      Left            =   7680
      Pattern         =   "*.kon"
      TabIndex        =   22
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "ì«çû"
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   480
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   4440
      Max             =   2000
      Min             =   10
      SmallChange     =   10
      TabIndex        =   17
      Top             =   3600
      Value           =   500
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   15.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Picture         =   "onngaku1.frx":03DE
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "çƒê∂"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ï€ë∂"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Left            =   7200
      Top             =   600
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   1
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   2
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   3
      Left            =   720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   4
      Left            =   840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   7
      Left            =   1560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   9
      Left            =   1920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   10
      Left            =   2160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   11
      Left            =   2280
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   12
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   13
      Left            =   2880
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   14
      Left            =   3000
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   15
      Left            =   3240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   16
      Left            =   3360
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   17
      Left            =   3600
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   18
      Left            =   3720
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   19
      Left            =   4080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   20
      Left            =   4320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   21
      Left            =   4440
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   22
      Left            =   4680
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   23
      Left            =   4800
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   24
      Left            =   5160
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   25
      Left            =   5400
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   26
      Left            =   5520
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   27
      Left            =   5760
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   28
      Left            =   5880
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   29
      Left            =   6120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   30
      Left            =   6240
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   930
      Index           =   31
      Left            =   6600
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1640
      _Version        =   393216
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\√ﬁΩ∏ƒØÃﬂ\å˜àÍ\My documents\âπ\do.wav"
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   29
      Left            =   6240
      TabIndex        =   73
      ToolTipText     =   "Å™ÅÚ◊"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   27
      Left            =   5880
      TabIndex        =   71
      ToolTipText     =   "Å™ÅÚø"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   25
      Left            =   5520
      TabIndex        =   69
      ToolTipText     =   "Å™ÅÚÃß"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   31
      Left            =   6720
      TabIndex        =   75
      ToolTipText     =   "Å™Å™ƒﬁ"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   30
      Left            =   6360
      TabIndex        =   74
      ToolTipText     =   "Å™º"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   28
      Left            =   6000
      TabIndex        =   72
      ToolTipText     =   "Å™◊"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   26
      Left            =   5640
      TabIndex        =   70
      ToolTipText     =   "Å™ø"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   24
      Left            =   5280
      TabIndex        =   68
      ToolTipText     =   "Å™Ãß"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   22
      Left            =   4800
      TabIndex        =   66
      ToolTipText     =   "Å™ÅÚ⁄"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   20
      Left            =   4440
      TabIndex        =   64
      ToolTipText     =   "Å™ÅÚƒﬁ"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   23
      Left            =   4920
      TabIndex        =   67
      ToolTipText     =   "Å™–"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   21
      Left            =   4560
      TabIndex        =   65
      ToolTipText     =   "Å™⁄"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   17
      Left            =   3720
      TabIndex        =   61
      ToolTipText     =   "ÅÚ◊"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   15
      Left            =   3360
      TabIndex        =   59
      ToolTipText     =   "ÅÚø"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   19
      Left            =   4200
      TabIndex        =   63
      ToolTipText     =   "Å™ƒﬁ"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   18
      Left            =   3840
      TabIndex        =   62
      ToolTipText     =   "º"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   16
      Left            =   3480
      TabIndex        =   60
      ToolTipText     =   "◊"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   13
      Left            =   3000
      TabIndex        =   57
      ToolTipText     =   "ÅÚÃß"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   14
      Left            =   3120
      TabIndex        =   58
      ToolTipText     =   "ø"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   10
      Left            =   2280
      TabIndex        =   54
      ToolTipText     =   "ÅÚ⁄"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   8
      Left            =   1920
      TabIndex        =   52
      ToolTipText     =   "ÅÚƒﬁ"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   12
      Left            =   2760
      TabIndex        =   56
      ToolTipText     =   "Ãß"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   11
      Left            =   2400
      TabIndex        =   55
      ToolTipText     =   "–"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   9
      Left            =   2040
      TabIndex        =   53
      ToolTipText     =   "⁄"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   7
      Left            =   1680
      TabIndex        =   51
      ToolTipText     =   "ƒﬁ"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   5
      Left            =   1200
      TabIndex        =   49
      ToolTipText     =   "Å´ÅÚ◊"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   6
      Left            =   1320
      TabIndex        =   50
      ToolTipText     =   "Å´º"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   3
      Left            =   840
      TabIndex        =   47
      ToolTipText     =   "Å´ÅÚø"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   45
      ToolTipText     =   "Å´ÅÚÃß"
      Top             =   2160
      Width           =   255
      BackColor       =   0
      Size            =   "450;1296"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   4
      Left            =   960
      TabIndex        =   48
      ToolTipText     =   "Å´◊"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   2
      Left            =   600
      TabIndex        =   46
      ToolTipText     =   "Å´ø"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Å´Ãß"
      Top             =   2160
      Width           =   375
      BackColor       =   16777215
      Size            =   "661;1931"
      FontName        =   "ÇlÇr ÇoÉSÉVÉbÉN"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   7200
      Picture         =   "onngaku1.frx":0528
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label3 
      Caption         =   "é©ìÆââët"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "ë¨Ç≥  ë¨Ç¢•••••••••••••••••••••••••••íxÇ¢"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "  O 0 P - @^  [  Z SX D C V GB HN J M Q 2 W3 E  R 5 T6 Y 7 U  I"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Menu mnfile 
      Caption         =   "ÉtÉ@ÉCÉã"
      Begin VB.Menu mnload 
         Caption         =   "ì«çû"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnsave 
         Caption         =   "ï€ë∂"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnb 
         Caption         =   "-"
      End
      Begin VB.Menu mnplay 
         Caption         =   "çƒê∂"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnstop 
         Caption         =   "àÍéûí‚é~"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnclear 
         Caption         =   "í‚é~"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnplay2 
         Caption         =   "ìríÜçƒê∂"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnb2 
         Caption         =   "-"
      End
      Begin VB.Menu mnend 
         Caption         =   "âπäy1èIóπ"
      End
   End
   Begin VB.Menu mn_omp 
      Caption         =   "âπïÑ"
      Tag             =   "0"
      Begin VB.Menu mn_omps 
         Caption         =   "ãxïÑ              Ctrl+0"
         Index           =   0
         Tag             =   "0"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "ëSâπïÑ           Ctrl+1"
         Index           =   1
         Tag             =   "16"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "ìÒï™âπïÑ        Ctrl+2"
         Index           =   2
         Tag             =   "8"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "ïtì_ìÒï™âπïÑ  Ctrl+3"
         Index           =   3
         Tag             =   "12"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "élï™âπïÑ        Ctrl+4"
         Index           =   4
         Tag             =   "4"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "ïtì_élï™âπïÑ  Ctrl+5"
         Index           =   5
         Tag             =   "6"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "î™ï™âπïÑ        Ctrl+8"
         Index           =   6
         Tag             =   "2"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "ïtì_î™ï™âπïÑ  Ctrl+9"
         Index           =   7
         Tag             =   "3"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "è\òZï™âπïÑ     Ctrl+6"
         Index           =   8
         Tag             =   "1"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "òAë±ì¸óÕ        Ctrl+-"
         Index           =   9
         Tag             =   "99"
      End
      Begin VB.Menu mn_omps 
         Caption         =   "é©ìÆãxïÑòAë±  Ctrl++"
         Index           =   10
         Tag             =   "98"
      End
      Begin VB.Menu mn_ompb1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_ompCancel 
         Caption         =   "∑¨›æŸ"
      End
   End
   Begin VB.Menu mnsets 
      Caption         =   "ê›íË"
      Begin VB.Menu mnset 
         Caption         =   "Ç±ÇÃã»ÇÃÃﬂ€ ﬂ√®..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnset2 
         Caption         =   "µÃﬂºÆ›..."
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnb3 
         Caption         =   "-"
      End
      Begin VB.Menu mnabout 
         Caption         =   "âπäy1Ç…ä÷Çµ..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ob(9999) As Byte, tlen As Double 'ob...old type b;tlen...time length
Dim nb(31, 9999) As Byte, tlen1 As Integer 'nb...new type b;tlen...time length
Dim bcnt As Integer, x1_a As Integer, x1_b 'bcnt...where play;x1_a...where design
Dim btndat As Byte, currentdir, nowtype As Integer
Dim omp_long(10) As Integer, omp_mode0 As Integer 'omp_mode0...design omp
Public omp_mode1 As Integer 'omp_mode1...shortest omp

'Commands
Private Sub Command2_Click()
If MsgBox("ï€ë∂ÇµÇƒÇ‡ó«Ç¢Ç≈Ç∑Ç©ÅH" & Chr(13) & "è„èëÇ´ï€ë∂Ç…íçà”ÇµÇƒÇ≠ÇæÇ≥Ç¢ÅI", 36, "ï€ë∂ämîF") = 7 Then MsgPanel "ï€ë∂ÇÕíÜé~Ç≥ÇÍÇ‹ÇµÇΩÅB": Exit Sub
If nowtype = 0 Then '1{
 Select Case MsgBox("Ç±ÇÃã»ÇÃì¸óÕï˚ñ@ÇämîFÇµÇ‹Ç∑ÅB" & Chr(13) & "Ç±ÇÃã»ÇÕtype2(òaâπëŒâû)ÇÃì¸óÕÇ≈Ç∑Ç©?", 35, "ã»ï€ë∂å`éÆämîF")
 Case 6: nowtype = 2
 Case 7: nowtype = 1
 Case 2: MsgPanel "ï€ë∂ÇÕíÜé~Ç≥ÇÍÇ‹ÇµÇΩÅB": Exit Sub
 End Select
End If '}1
Form1.MousePointer = 11
Open currentdir & Text2.Text For Output As 1
If nowtype = 2 Then '1{
 Print #1, "type2"
 For a = 0 To 31 '2{
  For b = 0 To tlen1 '3{
   c = c & nb(a, b)
  Next b '}3
  Print #1, c
  c = ""
 Next a '}2
 Print #1, HScroll1.Value & "," & tlen1 & "," & HScroll2.Value & "," & x1_a
 Print #1, omp_mode0 & "," & omp_mode1
ElseIf nowtype = 1 Then '1
 Print #1, Text1.Text
 Print #1, HScroll1.Value
End If '}1
Close #1
Form1.MousePointer = 99
MsgPanel currentdir & Text2.Text & "Ç…ï€ë∂ÇµÇ‹ÇµÇΩ!"
File1.Path = "c:\"
File1.Path = currentdir
End Sub

Private Sub Command3_Click()
Dim b
On Error GoTo ERR
Open currentdir & Text2.Text For Input As 1
Line Input #1, a
If a = "type2" Then
 nowtype = 2
 For a = 0 To 31
  Line Input #1, b
  ll = Len(b)
  For c = 1 To ll
   nb(a, c - 1) = Mid(b, c, 1)
  Next c
  For c = ll + 1 To 9999
   nb(a, c - 1) = 0
  Next c
 Next a
 Input #1, hscr1, tle, hscr2, x1_
 Close #1
 HScroll1.Value = hscr1
 tlen1 = tle
 HScroll2.Value = hscr2
 x1_a = x1_
 MsgPanel currentdir & Text2.Text & "Ç©ÇÁtype2Ç≈ì«Ç›çûÇ›Ç‹ÇµÇΩÅI"
 Call Picture1DRAW
Else
nowtype = 1
 Line Input #1, d
 Close #1
 Text1.Text = a
 tlen = Len(a)
 HScroll1.Value = d
 MsgPanel currentdir & Text2.Text & "Ç©ÇÁì«Ç›çûÇ›Ç‹ÇµÇΩÅI"
End If
Exit Sub
ERR:
MsgPanel ERR.Number & currentdir & Text2.Text & "Ç©ÇÁì«Ç›çûÇﬁÇÃÇ…é∏îsÇµÇ‹ÇµÇΩÅIê≥ÇµÇ¢ÉtÉ@ÉCÉãÇ©ämîFÇµÇƒÇ≠ÇæÇ≥Ç¢ÅI"
End Sub

Private Sub Command1_Click()
If nowtype = 0 Then nowtype = -5 + MsgBox("Ç±ÇÃã»ÇÃì¸óÕï˚ñ@ÇämîFÇµÇ‹Ç∑ÅB" & Chr(13) & "Ç±ÇÃã»ÇÕtype2(òaâπëŒâû)ÇÃì¸óÕÇ≈Ç∑Ç©?", 36, "ã»ì¸óÕå`éÆämîF")
If nowtype = 1 Then Timer1.Interval = HScroll1.Value Else Timer2.Interval = HScroll1.Value
Label3.BackColor = &HA00000
Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Command4_Click()
If nowtype = 0 Then nowtype = -5 + MsgBox("Ç±ÇÃã»ÇÃì¸óÕï˚ñ@ÇämîFÇµÇ‹Ç∑ÅB" & Chr(13) & "Ç±ÇÃã»ÇÕtype2(òaâπëŒâû)ÇÃì¸óÕÇ≈Ç∑Ç©?", 36, "ã»ì¸óÕå`éÆämîF")
If nowtype = 1 Then Timer1.Interval = 0 Else Timer2.Interval = 0
Label3.BackColor = &H8000000F
Label3.ForeColor = &H80000012
bcnt = 0
End Sub

Private Sub Command5_Click()
Dim sometimer As Timer
If nowtype = 0 Then nowtype = -5 + MsgBox("Ç±ÇÃã»ÇÃì¸óÕï˚ñ@ÇämîFÇµÇ‹Ç∑ÅB" & Chr(13) & "Ç±ÇÃã»ÇÕtype2(òaâπëŒâû)ÇÃì¸óÕÇ≈Ç∑Ç©?", 36, "ã»ì¸óÕå`éÆämîF")
If nowtype = 1 Then Set sometimer = Timer1 Else Set sometimer = Timer2
If Not sometimer.Interval = 0 Then
 sometimer.Interval = 0
 Label3.BackColor = &H8000000F
 Label3.ForeColor = &H80000012
ElseIf Not bcnt = 0 Then
 sometimer.Interval = HScroll1.Value
 Label3.BackColor = &HA00000
 Label3.ForeColor = &HFFFFFF
End If
End Sub

Private Sub Command6_Click()
If nowtype = 0 Then nowtype = -5 + MsgBox("Ç±ÇÃã»ÇÃì¸óÕï˚ñ@ÇämîFÇµÇ‹Ç∑ÅB" & Chr(13) & "Ç±ÇÃã»ÇÕtype2(òaâπëŒâû)ÇÃì¸óÕÇ≈Ç∑Ç©?", 36, "ã»ì¸óÕå`éÆämîF")
If nowtype = 1 Then Timer1.Interval = HScroll1.Value Else Timer2.Interval = HScroll1.Value
Label3.BackColor = &HA00000
Label3.ForeColor = &HFFFFFF
bcnt = x1_a
End Sub

'CommandButton1
Private Sub CommandButton1_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
Select Case KeyCode
Case 16
btndat = btndat + 1
Case 17
btndat = btndat + 2
Case 18
btndat = btndat + 4
End Select
End Sub

Private Sub CommandButton1_Click(Index As Integer)
Call Silent
MMControl1(Index).Command = "prev"
MMControl1(Index).Command = "play"
Call set_omp(Index, x1_a)
Call Picture1DRAW
End Sub

Private Sub CommandButton1_KeyUp(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim key As Integer
key = KeyCode
Call Picture1_KeyUp(key, Shift)
End Sub

Private Sub CommandButton1_LostFocus(Index As Integer)
btndat = 0
End Sub
'Filelist
Private Sub File1_Click()
Text2.Text = File1.List(File1.ListIndex)
End Sub
'Form
Private Sub Form_Load()
currentdir = CurDir
'******************exeèëèoÇÃéûÇ…éüÇÃçsÇ…Ç¬ÇØÇÈÅB********************'
currentdir = "C:\Documents and Settings\å˜àÍ\ÉfÉXÉNÉgÉbÉv\é©çÏÉvÉçÉOÉâÉÄ\âπäy\"
If Right(currentdir, 1) <> "\" Then currentdir = currentdir & "\"
File1.Path = currentdir
For a = 0 To 9999
 ob(a) = 255
Next a
Form2.Check4.Value = 1 '⁄
For a = 0 To 31
 MMControl1(a).Notify = False: MMControl1(a).Wait = True: MMControl1(a).Shareable = False: MMControl1(a).DeviceType = "WaveAudio"
Next a
Form2.Check1.Value = 1 '⁄
For a = 0 To 31
 MMControl1(a).FileName = currentdir & "wavdat\a" & a & ".wav"
Next a
Form2.Check2.Value = 1 '⁄
For a = 0 To 31
 Form2.ProgressBar1.Value = a: MMControl1(a).Command = "open"
Next a
For a = 0 To 32
 Picture1.Line (0, a * 120)-(Picture1.Width, a * 120)
Next a
For a = 0 To 62
 Picture1.Line (a * 120, 0)-(a * 120, Picture1.Height)
Next a
omp_mode1 = 2
For a = 0 To 10
omp_long(a) = mn_omps(a).Tag
Next a
Call mn_omps_Click(4)
Form2.Hide
Call Picture1DRAW
End Sub

Private Sub Form_Unload(Cancel As Integer)
For a = 0 To 31
MMControl1(a).Command = "close"
Next a
End
End Sub
'Scrolls
Private Sub HScroll2_Change()
Call Picture1DRAW
End Sub
'Menu

Private Sub mn_omps_Click(Index As Integer)
omp_mode = Index
End Sub

Private Sub mnabout_Click()
Form2.Show
End Sub

Private Sub mnclear_Click()
Call Command4_Click
End Sub

Private Sub mnend_Click()
End
End Sub

Private Sub mnload_Click()
Call Command3_Click
End Sub

Private Sub mnplay_Click()
Call Command1_Click
End Sub

Private Sub mnsave_Click()
Call Command2_Click
End Sub

Private Sub mnset_Click()
Form1.Enabled = False
Form3.Show
End Sub

Private Sub mnstop_Click()
Call Command5_Click
End Sub
'Picture1
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 16
If Not btndat Mod 2 = 1 Then btndat = btndat + 1
Case 17
If Not Int((btndat Mod 4) / 2) = 1 Then btndat = btndat + 2
Case 18
If Not Int(btndat / 4) = 1 Then btndat = btndat + 4
End Select
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
a = x1_a
Select Case KeyCode '1
Case 16: If btndat Mod 2 Then btndat = btndat - 1
Case 17: If Int(btndat / 2) Mod 2 Then btndat = btndat - 2
Case 18: If Int(btndat / 4) Then btndat = btndat - 4
Case Else
 Select Case btndat '2
 Case 0
  x1_a = x1_a + 1
  For gre = 0 To 31 '3
   If Mid("794880        90838868678671667278747781508751698253845489558573", 2 * gre + 1, 2) = KeyCode Then Call set_omp(gre, a)
  Next gre '3
  Select Case KeyCode '3
  Case 70
   For gr1 = 0 To 31 '4
    nb(gr1, a) = 2
   Next gr1 '4
  Case 93
   PopupMenu mn_omp
   x1_a = x1_a - 1
  Case 38
   Form1.omp_mode = Form1.omp_mode - 0.9
   x1_a = x1_a - 1
  Case 40
   Form1.omp_mode = Form1.omp_mode + 1
   x1_a = x1_a - 1
  Case 189: Call set_omp(3, a)
  Case 192: Call set_omp(4, a)
  Case 222: Call set_omp(5, a)
  Case 219: Call set_omp(6, a)
  Case 188, 37: x1_a = x1_a - 2
  End Select '3
 Case 2
  Select Case KeyCode - 48 '3
  Case 0 To 5: omp_mode = KeyCode - 48
  Case 8: omp_mode = 6
  Case 9: omp_mode = 7
  Case 6: omp_mode = 8
  Case 141: omp_mode = 9
  Case 139: omp_mode = 10
  End Select '3
 'Case 3: MsgPanel KeyCode
 End Select '2
End Select '1
If x1_a < 0 Then x1_a = 0
If x1_a > 9999 Then x1_a = 9999
If x1_a < HScroll2.Value Then HScroll2.Value = x1_a
If x1_a > HScroll2.Value + 61 Then HScroll2.Value = x1_a - 61
Call Picture1DRAW
End Sub

Private Sub Picture1_LostFocus()
btndat = 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
yy = Int(Y / 120)
Picture1.ToolTipText = CommandButton1(yy).ToolTipText
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Integer, yy As Integer
xx = Int(X / 120) + HScroll2.Value
yy = Int(Y / 120)
Select Case Button '1{
Case 1
 Call set_omp(yy, xx)
Case 2
 PopupMenu mn_omp
Case 4
 If Shift = 1 Then '2{
  tlen1 = xx
  MsgPanel tlen1 & "î‘ñ⁄ÇÃèÓïÒÇ‹Ç≈ââëtÇµÇ‹Ç∑ÅB"
  Beep
 Else
  nb(yy, xx) = 0
  x1_a = xx
  Picture1.Tag = ""
 End If '}2
End Select '}1
Call Picture1DRAW
End Sub
'Statusbar
Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
MsgBox "èÓïÒ:" & Panel.Text, vbInformation
End Sub
'Texts
Private Sub Text1_Change()
On Error GoTo err1
tlen = Len(Text1.Text)
On Error GoTo err2
For a = 1 To tlen '1
 alp = Mid(Text1.Text, a, 1)
 gr2 = 0
 For gre = 0 To 32 '2
  If Mid("o0p-@^[zsxdcvgbhnjmq2w3er5t6y7uif", gre + 1, 1) = alp Then
   ob(a) = gre
   gr2 = 1
  End If
 Next gre '2
 If gr2 = 0 Then ob(a) = 255
Next a '1
Call Picture1DRAW
Exit Sub
err1:
MsgPanel "ì¸óÕÇµÇΩï∂éöóÒÇì«Ç›çûÇﬁÇÃÇ…é∏îsÇµÇ‹ÇµÇΩÅI"
Exit Sub
err2:
MsgPanel a & "ï∂éöñ⁄ÇÃé©ìÆââëtópÉfÅ[É^ÇÃçÏê¨Ç…é∏îsÇµÇ‹ÇµÇΩÅI"
End Sub

Private Sub Text2_lostfocus()
If Right(Text2.Text, 4) <> ".kon" Then Text2.Text = Text2.Text & ".kon"
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then GoTo EXITS
alp = Text3.Text
Text3.Text = ""
a = x1_a
x1_a = x1_a + 1
For gre = 0 To 31
 If Mid("o0p-@^[zsxdcvgbhnjmq2w3er5t6y7ui", gre + 1, 1) = alp Or Mid("ZXCVBNMQSDGHJWERTYUI1#%&'OP`{=~", gre + 1, 1) = alp Then
  nb(gre, a) = 1
 End If
Next gre
Select Case alp
Case "f"
 For gr1 = 0 To 31
  nb(gr1, a) = 2
 Next gr1
Case ",", "<"
 x1_a = x1_a - 2
 If x1_a < 0 Then x1_a = 0
End Select
If x1_a < HScroll2.Value Then HScroll2.Value = x1_a
If x1_a > HScroll2.Value + 61 Then HScroll2.Value = x1_a - 61
Call Picture1DRAW
EXITS:
End Sub
'Timers
Private Sub Timer1_Timer()
If bcnt > tlen Then
Call Silent
Timer1.Interval = 0
Label3.BackColor = &H8000000F
Label3.ForeColor = &H80000012
bcnt = 0
Else
bcnt = bcnt + 1
If ob(bcnt) <> 255 Then
If ob(bcnt) = 32 Then
Call Silent
Else
Call Silent
MMControl1(ob(bcnt)).Command = "prev"
MMControl1(ob(bcnt)).Command = "play"
End If
End If
End If
End Sub

Private Sub Timer2_Timer()
If bcnt > tlen1 Then
Call Silent
Timer2.Interval = 0
Label3.BackColor = &H8000000F
Label3.ForeColor = &H80000012
bcnt = 0
Else
bcnt = bcnt + 1
For a = 0 To 31
If nb(a, bcnt) = 1 Then
If MMControl1(a).Mode = 526 Then MMControl1(a).Command = "stop"
MMControl1(a).Command = "prev"
MMControl1(a).Command = "play"
ElseIf nb(a, bcnt) = 2 Then
If MMControl1(a).Mode = 526 Then MMControl1(a).Command = "stop"
End If
Next a
End If
End Sub
'(General)
Public Sub MsgPanel(xxx)
StatusBar1.Panels(1).Text = xxx
StatusBar1.Panels(1).ToolTipText = xxx
Beep
End Sub

Public Sub Picture1DRAW()
For a = HScroll2.Value To HScroll2.Value + 61 '1
a1 = a - HScroll2.Value
If a > 9999 Then '2
For a0 = 0 To 3720 Step 120
a2 = a1 * 120
Picture1.Line (a2 + 15, a0 + 15)-(a2 + 105, a0 + 105), 0, BF
Next a0
Else '2
For a0 = 0 To 31 '3
a2 = a1 * 120
a3 = a0 * 120
Select Case nb(a0, a) '4
Case 0
If a = tlen1 Then
Picture1.Line (a2 + 15, a3 + 15)-(a2 + 105, a3 + 105), &HFFFF&, BF
Else
Picture1.Line (a2 + 15, a3 + 15)-(a2 + 105, a3 + 105), &HFFFFFF, BF
End If
Case 1
Picture1.Line (a2 + 15, a3 + 15)-(a2 + 105, a3 + 105), &HFF0000, BF
Case 2
Picture1.Line (a2 + 15, a3 + 15)-(a2 + 105, a3 + 105), &HFF&, BF
End Select '4
If a = x1_a Then Picture1.Line (a2 + 15, a3 + 15)-(a2 + 105, a3 + 105), &HFF00&, B
Next a0 '3
End If '2
Next a '1
End Sub

Public Sub Silent()
For a = 0 To 31
If MMControl1(a).Mode = 526 Then MMControl1(a).Command = "stop"
Next a
End Sub

Public Sub set_omp(yy, xx)
If Not (0 <= yy And yy <= 31 And 0 <= xx And xx <= 9999) Then Exit Sub
Select Case mn_omp.Tag
Case 0
 nb(yy, xx) = 2
 x1_a = xx + 1
Case 99
 nb(yy, xx) = 1
 x1_a = xx + 1
Case 98
 If Picture1.Tag <> "" Then nb(Picture1.Tag, xx) = 2
 Picture1.Tag = yy
 nb(yy, xx) = 1
 x1_a = xx + 1
Case Else '1*
 nb(yy, xx) = 1
 Dim mn0 As Integer: mn0 = mn_omp.Tag: mn0 = mn0 + xx
 For a1 = xx + 1 To mn0 '2{
  If a1 > 9999 Then '3{
   x1_a = 9999
   GoTo skip1 'Å®Åö
  ElseIf nb(yy, a1) = 1 Then '3*
   x1_a = mn0
   GoTo skip1 'Å®Åö
  ElseIf nb(yy, a) = 2 Then '3*
   nb(yy, a) = 0
  End If '}3
 Next a1 '}2
 nb(yy, mn0) = 2
 x1_a = mn0
skip1: 'Å©Åö
End Select
End Sub

Public Sub set_omp_long(Index As Integer, NewValue As Integer)
If Index < 0 Or Index > 10 Then Exit Sub
omp_long(Index) = NewValue
End Sub

Public Property Get omp_mode()
omp_mode = omp_mode0
End Property

Public Property Let omp_mode(ByVal NewValue)
Top:
Index = Int(NewValue)
If Index < 0 Then Index = 10
If Index > 10 Then Index = 0
If mn_omps(Index).Enabled = False Then
 If (NewValue * 10) Mod 10 = 0 Then NewValue = NewValue + 1 Else NewValue = NewValue - 1
 GoTo Top
End If
For a = 0 To 10
mn_omps(a).Checked = False
Next a
Picture1.Tag = ""
Form1.MouseIcon = LoadPicture(currentdir & "images\" & mn_omps(Index).Tag & ".cur")
mn_omp.Tag = omp_long(Index)
mn_omps(Index).Checked = True
omp_mode0 = Index
End Property
