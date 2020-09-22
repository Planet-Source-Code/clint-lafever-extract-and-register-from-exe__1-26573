VERSION 5.00
Object = "{14F718A0-FFF2-406A-9571-E944FCBDEB5D}#2.5#0"; "VBDice.ocx"
Begin VB.Form frmMAIN 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register Example"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdROLL 
      Caption         =   "&Roll"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VBDice.Dice dieDEFEND 
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      DiceColor       =   1
   End
   Begin VBDice.Dice dieATTACK 
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VBDice.Dice dieATTACK 
      Height          =   480
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VBDice.Dice dieATTACK 
      Height          =   480
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VBDice.Dice dieDEFEND 
      Height          =   480
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      DiceColor       =   1
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1800
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
'*******************************************
' READ READ READ READ READ READ READ
'*******************************************
'
' I am writting this here to let you know that
' when you first open this project, you will get
' an error saying it cannot find the VBDICE control.
'  The whole point of this program is to show you
' how a program can put some file onto a users
' machine and register it if needed.  So, if you
' trust me and my code, when you run it, a copy
' of VBDICE.OCX will be put on your machine and
' registered.  After that, because you are running
' this from the source, you will have to CLOSE
' VB WITHOUT saving.  Then reopen the project and
' all will work then because you have the control.
'  If you compile and give this to somebody [who
' has VB runtimes] when they run it, it will put
' the control on their computer and run perfectly.
'  Only because you are running this from the source
' are there problems.
'------------------------------------------------------------



Option Explicit
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Purpose:  Roll each dice a random number of times [1-15]
' Date: August,23 2001 @ 11:39:41
'------------------------------------------------------------
Private Sub cmdROLL_Click()
    On Error Resume Next
    Dim x As Long
    Randomize
    For x = 0 To Me.dieATTACK.UBound
        Me.dieATTACK(x).Roll CLng((15 * Rnd) + 1)
    Next x
    For x = 0 To Me.dieDEFEND.UBound
        Me.dieDEFEND(x).Roll CLng((15 * Rnd) + 1)
    Next x
End Sub
