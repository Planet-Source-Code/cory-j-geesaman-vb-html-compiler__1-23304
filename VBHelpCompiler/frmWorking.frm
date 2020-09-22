VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWorking 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compiling HTML Files"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   1320
      Top             =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "http://www.naven.net/"
      Default         =   -1  'True
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar tBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin SHDocVwCtl.WebBrowser wB 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Batch:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   465
   End
   Begin VB.Label txtFile 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "frmWorking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
wB.Navigate "http://www.naven.net/", -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Working = False
frmMain.lstSyntaxColors.Enabled = True
frmMain.Command1.Enabled = True
frmMain.Command2.Enabled = True
frmMain.Command3.Enabled = True
frmMain.Command4.Enabled = True
frmMain.Command5.Enabled = True
frmMain.Command6.Enabled = True
frmMain.Command7.Enabled = True
frmMain.cmdSyntaxColor.Enabled = True
frmMain.txtOutputDir.Enabled = True
frmMain.txtInputPRJ.Enabled = True
frmMain.Check1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If frmMain.Working = False Then
Unload Me
Else
Me.ZOrder 0
End If
End Sub
