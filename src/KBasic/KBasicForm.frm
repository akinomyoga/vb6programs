VERSION 5.00
Object = "{B30B7ED4-9187-4EC4-9CD3-5155839C07F7}#2.0#0"; "KBasic.ocx"
Begin VB.Form KBasicForm 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin KBasic.ColorButton ColorButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Color1"
      Top             =   1560
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      font            =   "KBasicForm.frx":0000
   End
   Begin KBasic.ToggleButton ToggleButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Toggle2"
      Top             =   1080
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      font            =   "KBasicForm.frx":002C
   End
   Begin KBasic.ToggleButton ToggleButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Toggle1"
      Top             =   600
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      backcolor       =   12648447
      forecolor       =   255
      font            =   "KBasicForm.frx":0058
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "KBasicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
