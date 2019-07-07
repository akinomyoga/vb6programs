VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "êwéÊÇËÉQÅ[ÉÄ"
   ClientHeight    =   7995
   ClientLeft      =   3420
   ClientTop       =   1710
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9900
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÉQÅ[ÉÄÇÃìríÜåoâﬂ"
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÉpÉX"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Player2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   8160
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Player1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "?-?"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu game 
      Caption         =   "ÉQÅ[ÉÄ"
      Begin VB.Menu start 
         Caption         =   "ÉXÉ^Å[Ég"
      End
   End
   Begin VB.Menu settei 
      Caption         =   "ê›íË"
      Begin VB.Menu cellcnt 
         Caption         =   "Ç‹Ç∑ÇÃêî..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ccnt As Integer, trn2 As Integer, nme As Integer
'trn2ÇÕÇ«ÇøÇÁÇÃÉvÉåÉCÉÑÅ[ÇÃêFÇåàÇﬂÇƒÇ¢ÇÈÇ©(Form3)
Public pl1cR As Integer, pl1cG As Integer, pl1cB As Integer
Public pl2cR As Integer, pl2cG As Integer, pl2cB As Integer
Dim cdat(29, 29) As Integer, trn As Integer, trnc, cdat2(29, 29) As Boolean, str(1 To 2) As Boolean
Dim plcnt(1 To 2) As Integer
Dim x0 As Integer, y0 As Integer, tcnt As Integer
Private Sub cellcnt_Click()
Form2.Show
End Sub

Private Sub Command1_Click()
Call NEXTPLAYER
End Sub

Private Sub Command2_Click()
For a = 0 To 1
b = b & Chr(13) & Label2(a).Caption & " " & plcnt(a + 1) & "Ç‹Ç∑"
Next a
MsgBox b
End Sub

Private Sub Command3_Click()
Call GAMEOVER
End Sub

Private Sub Form_Click()
If x0 <= ccnt - 1 And y0 <= ccnt - 1 And str(trn) = True Then
cdat2(x0, y0) = True
str(trn) = False
End If
If x0 <= ccnt - 1 And y0 <= ccnt - 1 And cdat2(x0, y0) = True Then
FillColor = trnc
Circle (x0 * 240 + 120, y0 * 240 + 120), 105, trnc
plcnt(trn) = plcnt(trn) + 1
cdat(x0, y0) = trn
tcnt = tcnt - 1
If tcnt <= 0 Then
Call NEXTPLAYER
End If
Call DRAWOK
Label4.Caption = "Ç†Ç∆" & tcnt & "âÒ"
Label5.Caption = Label2(trn - 1).Caption & "ÇÃî‘"
End If
End Sub

Private Sub Form_Load()
pl1cR = 0: pl1cG = 0: pl1cB = 0
pl2cR = 255: pl2cG = 255: pl2cB = 255
trn = 1
tcnt = SAIKORO
FillStyle = 0
ccnt = 30
Call DRAWCELL
nme = 6
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
x0 = Int(X / 240)
y0 = Int(Y / 240)
If x0 <= ccnt - 1 And y0 <= ccnt - 1 Then
Label1.Caption = x0 & "-" & y0
Else
Label1.Caption = "?-?"
End If
End Sub

Public Sub DRAWCELL()
Form1.Cls
a2 = 240 * ccnt
Form1.DrawWidth = 1
For a = 0 To ccnt
a1 = a * 240
Form1.Line (a1, a2)-(a1, 0), RGB(0, 0, 0)
Form1.Line (a2, a1)-(0, a1), RGB(0, 0, 0)
Next a
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_Click()
trn2 = 1
With Form3
.Picture1.BackColor = Picture1.BackColor
.HScroll1.Value = pl1cR
.HScroll2.Value = pl1cG
.HScroll3.Value = pl1cB
.Caption = "Player1ÇÃêF"
.Show
End With
End Sub

Public Function NEWTRNC()
If trn = 1 Then NEWTRNC = Picture1.BackColor
If trn = 2 Then NEWTRNC = Picture2.BackColor
End Function

Private Sub Picture2_Click()
trn2 = 2
With Form3
.Picture1.BackColor = Picture2.BackColor
.HScroll1.Value = pl2cR
.HScroll2.Value = pl2cG
.HScroll3.Value = pl2cB
.Caption = "Player2ÇÃêF"
.Show
End With
End Sub

Public Function SAIKORO()
Randomize
a = Int(Rnd * nme)
If a = 0 Then a = nme
SAIKORO = a
End Function

Public Sub DRAWSTONE()
a = FillColor
For X = 0 To ccnt - 1
For Y = 0 To ccnt - 1
If cdat(X, Y) = 1 Then
FillColor = RGB(pl1cR, pl1cG, pl1cB)
Circle (X * 240 + 120, Y * 240 + 120), 105, FillColor
End If
If cdat(X, Y) = 2 Then
FillColor = RGB(pl2cR, pl2cG, pl2cB)
Circle (X * 240 + 120, Y * 240 + 120), 105, FillColor
End If
Next Y
Next X
FillColor = a
End Sub

Public Sub DRAWOK()
Dim okcnt As Boolean
FillStyle = 1
okcnt = True
For X = 0 To ccnt - 1
For Y = 0 To ccnt - 1
cdat2(X, Y) = False
If X > 0 Then
If cdat(X - 1, Y) = trn Then GoTo OK
End If
If X < ccnt - 1 Then
If cdat(X + 1, Y) = trn Then GoTo OK
End If
If Y > 0 Then
If cdat(X, Y - 1) = trn Then GoTo OK
End If
If Y < ccnt - 1 Then
If cdat(X, Y + 1) = trn Then GoTo OK
End If
GoTo SKIP
OK:
If cdat(X, Y) = 0 Then
Circle (X * 240 + 120, Y * 240 + 120), 105, trnc
cdat2(X, Y) = False
okcnt = False
End If
SKIP:
Next Y
Next X
FillStyle = 0
If okcnt = 0 Then Call GAMEOVER
End Sub

Private Sub start_Click()
str(1) = True
str(2) = True
Form4.Show
End Sub

Public Sub NEXTPLAYER()
tcnt = SAIKORO
trn = NXTRN
trnc = NEWTRNC
Call DRAWCELL
Call DRAWSTONE
Call DRAWOK
MsgBox "éüÇÕ" & Label2(trn - 1).Caption & "ÇÃî‘Ç≈Ç∑ÅB"
End Sub

Public Function NXTRN()
If trn = 1 Then
NXTRN = 2
Else
NXTRN = 1
End If
End Function

Public Sub GAMEOVER()
a = NXTRN
plcnt(1) = 0
plcnt(2) = 0
For X = 0 To ccnt - 1
For Y = 0 To ccnt - 1
If cdat(X, Y) = 0 Then cdat(X, Y) = a
b = cdat(X, Y)
plcnt(b) = plcnt(b) + 1
Next Y
Next X
c = "   ***åãâ ***"
For d = 1 To 2
c = c & Chr(13) & "   *" & Label2(d - 1).Caption & " : " & plcnt(d) & "ì_*"
Next d
If plcnt(1) > plcnt(2) Then
c = c & Chr(13) & "   {" & Label2(1).Caption & "ÇÃèüÇøÅI}"
ElseIf plcnt(1) < plcnt(2) Then
c = c & Chr(13) & "   {" & Label2(2).Caption & "ÇÃèüÇøÅI}"
Else
c = c & Chr(13) & "   {à¯Ç´ï™ÇØ}"
End If
MsgBox c
End Sub
