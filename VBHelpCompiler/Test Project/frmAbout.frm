VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About NSocks"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wB 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1080
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
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.naven.net/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1500
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   $"frmAbout.frx":0152
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Image sP 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   240
      Picture         =   "frmAbout.frx":0266
      Top             =   240
      Width           =   390
   End
   Begin VB.Image cP 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   240
      Picture         =   "frmAbout.frx":0968
      Top             =   240
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NSocks"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   825
      TabIndex        =   1
      Top             =   225
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NSocks"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   480
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OverLabel As Boolean

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If App.Minor < 10 And App.Revision < 10 Then
Label3.Caption = "v" & App.Major & "." & App.Minor & App.Revision
Else
Label3.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End If
Me.Caption = Me.Caption & " - " & Label3.Caption
End Sub

Public Sub InitAboutBox(Client As Boolean)
If Client = True Then
sP.Visible = False
cP.Visible = True
Label1.Caption = "NSocks - Client"
Else
cP.Visible = False
sP.Visible = True
Label1.Caption = "NSocks - Server"
End If
Label2.Caption = Label1.Caption
Me.ZOrder 0
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X < 0 Or Y < 0 Or X > Label5.Width Or Y > Label5.Height Then
OverLabel = False
Label5.ForeColor = &HC0&
Else
OverLabel = True
Label5.ForeColor = &HC00000
End If
End If
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And OverLabel = True Then
Label5.ForeColor = &HC0&
wB.Navigate Label5.Caption, -1
End If
End Sub
