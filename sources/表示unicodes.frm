VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo err
CommonDialog1.ShowOpen
a = CommonDialog1.filename
Open a For Input As 1
Open "c:\windows\ﾃﾞｽｸﾄｯﾌﾟ\unicodes出力ファイル.txt" For Output As 2
Do While Not EOF(1)
Line Input #1, b
For ee = 1 To Len(b)
If Len(b) > 0 Then
c = Asc(Mid(b, ee, 1))
d = AscW(Mid(b, ee, 1))
Print #2, c & ":" & d
End If
Next ee
Loop
Close #2
Close #1
err:
End
End Sub
