VERSION 5.00
Begin VB.UserControl KProgressBar 
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000002&
   ScaleHeight     =   1230
   ScaleWidth      =   1740
   Tag             =   "0"
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         ScaleHeight     =   735
         ScaleWidth      =   735
         TabIndex        =   1
         Top             =   0
         Width           =   735
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   225
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   210
      End
   End
End
Attribute VB_Name = "KProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' KProgressBar

Dim m_Value As Long

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal new_Value As Long)
    If m_Value <> new_Value Then
        m_Value = new_Value
        PropertyChanged "Value"
        update_state
    End If
End Property

Private Sub update_state()
    Picture1.BackColor = UserControl.BackColor
    Label1.BackColor = UserControl.BackColor
    Picture2.BackColor = UserControl.ForeColor
    Label2.BackColor = UserControl.ForeColor

    Dim label_text As String
    label_text = m_Value & "%"
    Label1.Caption = label_text
    Label1.Font = UserControl.Font
    Label2.Caption = label_text
    Label2.Font = UserControl.Font

    Dim label_x As Long, label_y As Long
    label_x = (Picture1.Width - Label1.Width) / 2
    label_y = (Picture1.Height - Label1.Height) / 2

    Label1.Left = label_x
    Label1.Top = label_y
    Label2.Left = label_x
    Label2.Top = label_y

    Picture2.Height = Picture1.Height
    Picture2.Width = Picture1.Width * m_Value / 100
    Picture2.Left = 0
    Picture2.Top = 0
End Sub

Private Sub update_size()
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    Picture1.Left = 0
    Picture1.Top = 0
    update_state
End Sub

Private Sub UserControl_Paint()
    update_state
End Sub

Private Sub UserControl_Resize()
    update_size
End Sub

Private Sub UserControl_Show()
    update_size
End Sub
