VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  'å≈íË¬∞Ÿ ≥®›ƒﬁ≥
   Caption         =   "ê›íË"
   ClientHeight    =   4635
   ClientLeft      =   3120
   ClientTop       =   2025
   ClientWidth     =   5610
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "êFÇÃê›íË"
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Ã◊Øƒ
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   3120
         ScaleHeight     =   1215
         ScaleWidth      =   1335
         TabIndex        =   44
         Top             =   1560
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   135
         Index           =   2
         Left            =   2520
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   47
         Top             =   2175
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'âEëµÇ¶
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   1800
         TabIndex        =   49
         Text            =   "0"
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   135
         Index           =   1
         Left            =   2520
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   48
         Top             =   1935
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'âEëµÇ¶
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1800
         TabIndex        =   46
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   135
         Index           =   0
         Left            =   2520
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   45
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'âEëµÇ¶
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   1800
         TabIndex        =   43
         Text            =   "0"
         Top             =   1665
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   33
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "CMY êFóøéOå¥êF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   1455
         Begin VB.PictureBox Picture2 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   39
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Index           =   4
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   38
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFF00&
            Height          =   255
            Index           =   3
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Text            =   "255"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Text            =   "255"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Text            =   "255"
            Top             =   720
            Width           =   615
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   3
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Width           =   255
            ForeColor       =   16776960
            BackColor       =   8421376
            Size            =   "450;370"
            Max             =   255
            Position        =   255
            Orientation     =   1
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   4
            Left            =   720
            TabIndex        =   30
            Top             =   480
            Width           =   255
            ForeColor       =   16711935
            BackColor       =   8388736
            Size            =   "450;370"
            Max             =   255
            Position        =   255
            Orientation     =   1
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   5
            Left            =   720
            TabIndex        =   29
            Top             =   720
            Width           =   255
            ForeColor       =   65535
            BackColor       =   32896
            Size            =   "450;370"
            Max             =   255
            Position        =   255
            Orientation     =   1
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "RGB êFåıéOå¥êF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1455
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   36
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   35
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   315
            TabIndex        =   34
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'âEëµÇ¶
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   1
            Left            =   720
            TabIndex        =   24
            Top             =   480
            Width           =   255
            ForeColor       =   49152
            BackColor       =   12648384
            Size            =   "450;370"
            Max             =   255
            Position        =   1
            Orientation     =   1
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   0
            Left            =   720
            TabIndex        =   23
            Top             =   240
            Width           =   255
            ForeColor       =   255
            BackColor       =   12632319
            Size            =   "450;370"
            Max             =   255
            Position        =   1
            Orientation     =   1
         End
         Begin MSForms.SpinButton SpinButton3 
            Height          =   210
            Index           =   2
            Left            =   720
            TabIndex        =   22
            Top             =   720
            Width           =   255
            ForeColor       =   16711680
            BackColor       =   16761024
            Size            =   "450;370"
            Max             =   255
            Position        =   1
            Orientation     =   1
         End
      End
      Begin MSForms.SpinButton SpinButton4 
         Height          =   135
         Index           =   2
         Left            =   2280
         TabIndex        =   53
         Top             =   2160
         Width           =   255
         Size            =   "450;238"
         Min             =   -1
         Max             =   1536
      End
      Begin MSForms.SpinButton SpinButton4 
         Height          =   135
         Index           =   1
         Left            =   2280
         TabIndex        =   52
         Top             =   1920
         Width           =   255
         Size            =   "450;238"
         Max             =   255
      End
      Begin MSForms.SpinButton SpinButton4 
         Height          =   135
         Index           =   0
         Left            =   2280
         TabIndex        =   51
         Top             =   1680
         Width           =   255
         Size            =   "450;238"
         Max             =   255
      End
      Begin VB.Label Label4 
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
         Height          =   135
         Left            =   1440
         TabIndex        =   32
         Top             =   1350
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1335
         X2              =   1335
         Y1              =   120
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1320
         X2              =   1320
         Y1              =   120
         Y2              =   2880
      End
      Begin VB.Label Label5 
         Caption         =   "êFëä"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   1440
         TabIndex        =   42
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "ñæìx"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   1440
         TabIndex        =   41
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "ç ìx"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   1440
         TabIndex        =   40
         Top             =   1680
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "å`"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton Command1 
         Caption         =   "ï“èW..."
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ï“èW"
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
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ïHå`"
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
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "í∑ï˚å`"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "â~"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ëÂÇ´Ç≥-íPà  Àﬂ∏æŸ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      TabIndex        =   7
      Top             =   360
      Width           =   2055
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Text            =   "1"
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Å´"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "â°ÇècÇ…ëµÇ¶ÇÈ"
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Å™"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "ècÇâ°Ç…ëµÇ¶ÇÈ"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   135
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin MSForms.SpinButton SpinButton2 
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   255
         Size            =   "450;344"
         Min             =   1
         Position        =   1
         Orientation     =   1
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   210
         Width           =   255
         Size            =   "450;344"
         Min             =   1
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label Label2 
         Caption         =   "â°ÇÃëÂÇ´Ç≥"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   12
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ècÇÃëÂÇ´Ç≥"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   1058
      TabFixedHeight  =   370
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "å`ëæÇ≥"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "êF"
            Key             =   "colortab"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
SpinButton1.Value = SpinButton2.Value
End Sub

Private Sub Command3_Click()
SpinButton2.Value = SpinButton1.Value
End Sub

Private Sub Form_Load()
e = Picture4.Height / 2
f = Picture4.Width / 2
Picture4.Cls
Dim c(5), d(5), r(5), g(5), bl(5)
For a = 0 To 63
For b = 0 To 63
r(0) = 255 / 63 * b
g(0) = a * 4 / 63 * b
bl(0) = 0
r(1) = (63 - a) * 4 / 63 * b
g(1) = 255 / 63 * b
bl(1) = 0
r(2) = 0
g(2) = 255 / 63 * b
bl(2) = a * 4 / 63 * b
r(3) = 0
g(3) = (63 - a) * 4 / 63 * b
bl(3) = 255 / 63 * b
r(4) = a * 4 / 63 * b
g(4) = 0
bl(4) = 255 / 63 * b
r(5) = 255 / 63 * b
g(5) = 0
bl(5) = (63 - a) * 4 / 63 * b
For h = 0 To 5
c(h) = cos1(a, b, h)
d(h) = sin1(a, b, h)
Picture4.PSet (-c(h) + f, -d(h) + e), RGB(r(h), g(h), bl(h))
Next h
Next b
Next a
End Sub

Private Sub SpinButton1_Change()
Text1.Text = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
Text2.Text = SpinButton2.Value
End Sub

Private Sub SpinButton3_Change(Index As Integer)
Text3(Index).Text = SpinButton3(Index).Value
If Index < 3 Then
SpinButton3(Index + 3).Value = 255 - SpinButton3(Index).Value
Picture2(Index).BackColor = SpinButton3(Index).Value * 256 ^ Index

Dim b As Integer, c As Integer, dd(2) As Integer, e As Integer
b = 0: c = 255
For a = 0 To 2
dd(a) = Text3(a).Text
If b < dd(a) Then
bb = a
b = dd(a)
End If
If c > dd(a) Then
cc = a
c = dd(a)
End If
e = e + dd(a)
Next a
SpinButton4(0).Value = b - c
SpinButton4(1).Value = e / 3
For a = 0 To 2
If a <> bb And a <> cc Then aa = a
Next a
aa1 = Int(Text3(aa).Text * 255 / Text3(bb).Text)
Select Case bb
Case 0
If aa = 1 Then
aa2 = aa1
Else
aa2 = 1535 - aa1
End If
Case 1
If aa = 0 Then
aa2 = 511 - aa1
Else
aa2 = 511 + aa1
End If
Case 2
If aa = 0 Then
aa2 = 1023 + aa1
Else
aa2 = 1023 - aa1
End If
End Select
SpinButton4(2).Value = aa2

Else
SpinButton3(Index - 3).Value = 255 - SpinButton3(Index).Value
Picture2(Index).BackColor = RGB(255, 255, 255) - SpinButton3(Index).Value * 256 ^ (Index - 3)
End If
Picture1.BackColor = RGB(SpinButton3(0).Value, SpinButton3(1).Value, SpinButton3(2).Value)
End Sub

Private Sub SpinButton4_Change(Index As Integer)
Select Case Index
Case 0
Case 1
Picture3(1).BackColor = &H10101 * SpinButton4(1).Value
Case 2
Dim spin2 As Integer: spin2 = SpinButton4(2).Value
Select Case spin2
Case -1: SpinButton4(2).Value = 1535
Case 0 To 255: Picture3(2).BackColor = RGB(255, spin2, 0)
Case 256 To 511: Picture3(2).BackColor = RGB(511 - spin2, 255, 0)
Case 512 To 767: Picture3(2).BackColor = RGB(0, 255, spin2 - 512)
Case 768 To 1023: Picture3(2).BackColor = RGB(0, 1023 - spin2, 255)
Case 1024 To 1279: Picture3(2).BackColor = RGB(spin2 - 1024, 0, 255)
Case 1280 To 1535: Picture3(2).BackColor = RGB(255, 0, 1535 - spin2)
Case 1536: SpinButton4(2).Value = 0
End Select
End Select
Text4(Index).Text = SpinButton4(Index).Value
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Frame1.Visible = True
Frame2.Visible = True
Frame5.Visible = False
Case 2
Frame1.Visible = False
Frame2.Visible = False
Frame5.Visible = True
End Select
End Sub

Private Sub Text1_Change()
On Error GoTo err1
SpinButton1.Value = Text1.Text
Exit Sub
err1:
Beep
Text1.Text = SpinButton1.Value
Label3.Caption = "1Ç©ÇÁ100ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢!"
End Sub

Private Sub Text2_Change()
On Error GoTo err1
SpinButton2.Value = Text2.Text
Exit Sub
err1:
Beep
Text2.Text = SpinButton2.Value
Label3.Caption = "1Ç©ÇÁ100ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢!"
End Sub

Private Sub Text3_Change(Index As Integer)
On Error GoTo err1
SpinButton3(Index).Value = Text3(Index).Text
Exit Sub
err1:
Beep
Label4.Caption = "0Ç©ÇÁ255ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢!"
End Sub

Private Sub Text3_LostFocus(Index As Integer)
On Error GoTo err1
SpinButton3(Index).Value = Text3(Index).Text
Exit Sub
err1:
Text3(Index).Text = SpinButton3(Index).Value
Beep
Label4.Caption = "0Ç©ÇÁ255ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢!"
End Sub

Private Sub Text4_Change(Index As Integer)
On Error GoTo err1
SpinButton4(Index).Value = Text4(Index).Text
Exit Sub
err1:
Beep
If Index = 2 Then
Label3.Caption = "0Ç©ÇÁ1535ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢"
Else
Label3.Caption = "0Ç©ÇÁ255ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢"
End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
On Error GoTo err1
SpinButton4(Index).Value = Text4(Index).Text
Exit Sub
err1:
Text4(Index) = SpinButton4(Index).Value
Beep
If Index = 2 Then
Label3.Caption = "0Ç©ÇÁ1535ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢"
Else
Label3.Caption = "0Ç©ÇÁ255ñòÇÃêîéöÇ≈ì¸óÕÇµÇƒâ∫Ç≥Ç¢"
End If
End Sub

Public Function cos1(a, b, c)
cos1 = Cos((a / 191 + c / 3) * 3.14159265359) * b / 8 * 75
End Function

Public Function sin1(a, b, c)
sin1 = Sin((a / 191 + c / 3) * 3.14159265359) * b / 8 * 75
End Function
