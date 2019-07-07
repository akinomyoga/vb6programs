VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "É{Å[ÉãÇ§Çø"
   ClientHeight    =   6765
   ClientLeft      =   3150
   ClientTop       =   2640
   ClientWidth     =   10575
   Icon            =   "É{Å[ÉãÇ§Çø.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10575
   Begin VB.VScrollBar VScroll3 
      Height          =   3135
      Left            =   10200
      Max             =   100
      TabIndex        =   11
      Top             =   2880
      Value           =   100
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3135
      Left            =   9960
      Max             =   0
      Min             =   100
      TabIndex        =   9
      Top             =   2880
      Value           =   100
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3135
      Left            =   9720
      Max             =   0
      Min             =   100
      TabIndex        =   7
      Top             =   2880
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   6120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ñﬂÇ∑"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ë≈Ç¬"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "èIóπ"
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   1000
      TabIndex        =   1
      Top             =   6120
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   0
      Min             =   90
      TabIndex        =   0
      Top             =   6360
      Value           =   90
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "ãÛãCíÔçR"
      Height          =   735
      Left            =   10200
      TabIndex        =   12
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "íµÇÀÇÈó "
      Height          =   735
      Left            =   9960
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "èdóÕ"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'ìhÇËÇ¬Ç‘Çµ
      Height          =   255
      Left            =   0
      Shape           =   3  'â~
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "äpìx"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Power"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   6120
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   9585
      X2              =   0
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'ìhÇËÇ¬Ç‘Çµ
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim b
Dim c
Dim d
Private Sub Command1_Click()
End
End Sub
Private Sub Command2_Click()
Command2.Enabled = False
Command3.Enabled = True
c = a
d = b
Timer1.Interval = 100
End Sub
Private Sub Command3_Click()
Command2.Enabled = True
Command3.Enabled = False
c = 0
d = 0
Shape1.Top = 5760
Shape1.Left = 0
Timer1.Interval = 0
End Sub
Private Sub HScroll1_Change()
Call HScroll2_Change
End Sub
Private Sub HScroll2_Change()
a = Int(HScroll2.Value * Sin(HScroll1.Value * 3.1415926535 / 180))
b = Int(HScroll2.Value * Cos(HScroll1.Value * 3.1415926535 / 180))
End Sub
Private Sub Timer1_Timer()
Shape1.Top = Shape1.Top - c
Shape1.Left = Shape1.Left + d
c = Int((c - VScroll1.Value) * VScroll3.Value / 100)
d = Int(d * VScroll3.Value / 100)
If Shape1.Top < 0 Then
Shape1.Top = 0
c = c * VScroll2.Value / -100
End If
If Shape1.Top > 5760 Then
c = c * VScroll2.Value / -100
Shape1.Top = 5760
End If
If Shape1.Left < 0 Then
Shape1.Left = 0
d = d * VScroll2.Value / -100
End If
If Shape1.Left > 9360 Then
Shape1.Left = 9360
d = d * VScroll2.Value / -100
End If
End Sub
Private Sub VScroll1_Change()
Line1.BorderColor = RGB(0, Int(VScroll1.Value * 2.55), 0)
End Sub
Private Sub VScroll2_Change()
Shape1.FillColor = RGB(155 + VScroll2.Value, 0, 0)
End Sub
Private Sub VScroll3_Change()
Shape2.FillColor = RGB(0, 0, Int(255 - VScroll3.Value * 2.55))
End Sub
