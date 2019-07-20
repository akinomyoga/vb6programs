VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#9.0#0"; "KBasic.ocx"
Begin VB.UserControl KColor 
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   2160
   ScaleWidth      =   4095
   Begin KBasic.SpinButton SpinRGB 
      Height          =   255
      Index           =   2
      Left            =   3240
      Top             =   720
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   16711680
      Max             =   255
   End
   Begin VB.TextBox TextH 
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Text            =   "00"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox TextD 
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
   Begin KBasic.ScrollBar ScrollRGB 
      Height          =   255
      Index           =   2
      Left            =   1080
      Top             =   720
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   16711680
      Max             =   255
      SmallChange     =   5
      BarSize         =   9
      LargeChange     =   10
   End
   Begin KBasic.SpinButton SpinRGB 
      Height          =   255
      Index           =   1
      Left            =   3240
      Top             =   480
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   49152
      Max             =   255
   End
   Begin VB.TextBox TextH 
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Text            =   "00"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TextD 
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin KBasic.ScrollBar ScrollRGB 
      Height          =   255
      Index           =   1
      Left            =   1080
      Top             =   480
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   49152
      Max             =   255
      SmallChange     =   5
      BarSize         =   9
      LargeChange     =   10
   End
   Begin KBasic.SpinButton SpinRGB 
      Height          =   255
      Index           =   0
      Left            =   3240
      Top             =   240
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   255
      Max             =   255
   End
   Begin VB.TextBox TextH 
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Text            =   "00"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox TextD 
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   375
   End
   Begin KBasic.ScrollBar ScrollRGB 
      Height          =   255
      Index           =   0
      Left            =   1080
      Top             =   240
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      ForeColor       =   255
      Max             =   255
      SmallChange     =   5
      BarSize         =   9
      LargeChange     =   10
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "#000000"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "KColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_currentControl As Control
Dim m_rgbValue(3) As Integer

Private Declare Function GetSysColor Lib "User32.dll" (ByVal nIndex As Long) As Long
 
Public Property Get Color() As OLE_COLOR
    Color = RGB(m_rgbValue(0), m_rgbValue(1), m_rgbValue(2))
End Property

Public Property Let Color(ByVal new_Color As OLE_COLOR)
    If Picture1.BackColor = new_Color Then Exit Property
    Picture1.BackColor = new_Color
    If (new_Color And &H80000000) <> 0 Then
        new_Color = GetSysColor(new_Color)
    End If
    setRGBComponent 0, new_Color And &HFF&
    setRGBComponent 1, (new_Color \ &H80&) And &HFF&
    setRGBComponent 2, (new_Color \ &H8000&) And &HFF&
End Property

Private Function Hex2(ByVal Value As Integer) As String
    Hex2 = Right("0" & Hex$(Value), 2)
End Function

''-----------------------------------------------------------------------------
''
'' èàóù
''
''-----------------------------------------------------------------------------

Private Sub setCurrentControl(ByRef new_currentControl As Control)
    Set m_currentControl = new_currentControl
End Sub

Private Sub resetCurrentControl(ByRef old_currentControl As Control)
    If m_currentControl Is old_currentControl Then
        Set m_currentControl = Nothing
    End If
End Sub

Private Sub updateLabel()
    Label1.Caption = "#" & Hex2(m_rgbValue(0)) & Hex2(m_rgbValue(1)) & Hex2(m_rgbValue(2))
    Picture1.BackColor = Me.Color
End Sub

Private Sub setRGBComponent(ByVal Index As Integer, ByVal Value As Integer)
    If Value < 0 Then
        Value = 0
    ElseIf Value > 255 Then
        Value = 255
    End If
    If m_rgbValue(Index) = Value Then Exit Sub

    m_rgbValue(Index) = Value
    ScrollRGB(Index).Value = Value
    SpinRGB(Index).Value = Value
    If Not (m_currentControl Is TextD(Index)) Then TextD(Index).Text = Value
    If Not (m_currentControl Is TextH(Index)) Then TextH(Index).Text = Hex2(Value)
    updateLabel
End Sub

''-----------------------------------------------------------------------------
''
'' ÉCÉxÉìÉgìoò^
''
''-----------------------------------------------------------------------------

Private Sub ScrollRGB_Scroll(Index As Integer)
    setRGBComponent Index, ScrollRGB(Index).Value
End Sub

Private Sub SpinRGB_Change(Index As Integer)
    setRGBComponent Index, SpinRGB(Index).Value
End Sub

Private Sub TextD_Change(Index As Integer)
    Dim l_value As Integer
    On Error GoTo label_IgnoreError
    l_value = CInt(TextD(Index).Text)
    On Error GoTo 0
    
    setRGBComponent Index, l_value
label_IgnoreError:
End Sub

Private Sub TextD_GotFocus(Index As Integer)
    setCurrentControl TextD(Index)
End Sub

Private Sub TextD_LostFocus(Index As Integer)
    resetCurrentControl TextD(Index)
    TextD(Index).Text = m_rgbValue(Index)
End Sub

Private Sub TextH_Change(Index As Integer)
    Dim l_value As Integer
    On Error GoTo label_IgnoreError
    l_value = CInt("&H" & TextH(Index).Text)
    On Error GoTo 0
    
    setRGBComponent Index, l_value
label_IgnoreError:
End Sub

Private Sub TextH_GotFocus(Index As Integer)
    setCurrentControl TextH(Index)
End Sub

Private Sub TextH_LostFocus(Index As Integer)
    resetCurrentControl TextH(Index)
    TextH(Index).Text = Hex2(m_rgbValue(Index))
End Sub
