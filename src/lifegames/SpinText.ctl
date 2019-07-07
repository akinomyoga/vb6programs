VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl SpinText 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   1605
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Text            =   "0"
      Top             =   0
      Width           =   975
   End
   Begin MSComCtl2.UpDown Spin1 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "SpinText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Property Get SpinMax() As Integer
    SpinMax = Spin1.Max
End Property

Public Property Let SpinMax(ByVal vNewValue As Integer)
    Spin1.Max = vNewValue
End Property

Public Property Get SpinMin() As Integer
    SpinMin = Spin1.Min
End Property

Public Property Let SpinMin(ByVal vNewValue As Integer)
    Spin1.Min = vNewValue
End Property

Public Property Get LabelCaption() As Variant
    LabelCaption = Label1.Caption
End Property

Public Property Let LabelCaption(ByVal vNewValue As Variant)
    Label1.Caption = vNewValue
End Property

Public Property Get LabelFont() As Variant
    LabelFont = Label1.Font
End Property

Public Property Let LabelFont(ByVal vNew As Variant)
    Label1.Font = vNew
End Property

Public Property Get TextAlign() As Integer
    TextAlign = Text1.Alignment
End Property

Public Property Let TextAlign(ByVal vNew As Integer)
    Text1.Alignment = vNew
End Property

Public Property Get SpinValue() As Integer
    SpinValue = Spin1.Value
End Property

Public Property Let SpinValue(ByVal vNew As Integer)
    Spin1.Value = vNew
    Text1.Text = vNew
End Property

Public Property Get LabelWidth() As Integer
    LabelWidth = Label1.Width
End Property

Public Property Let LabelWidth(ByVal vNew As Integer)
    Label1.Width = vNew
    Call LeftA
End Property

Public Property Get TextWidth() As Integer
    TextWidth = Text1.Width
End Property

Public Property Let TextWidth(ByVal vNew As Integer)
    Text1.Width = vNew
    Call LeftA
End Property

Private Sub Spin1_Change()
    Text1.Text = Spin1.Value
End Sub

Private Sub Text1_Change()
    On Error GoTo err
    Spin1.Value = Text1.Text
    Exit Sub
err:
    Text1.Text = Spin1.Value
End Sub

Public Sub LeftA()
    Text1.Left = Label1.Width + 105
    Spin1.Left = Text1.Left + Text1.Width - 15
End Sub
