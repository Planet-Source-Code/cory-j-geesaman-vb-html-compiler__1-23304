VERSION 5.00
Object = "*\A..\Project1.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin prjNSocks.Client Client 
      Height          =   390
      Left            =   4920
      TabIndex        =   7
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send To &Client"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send To &Server"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox TextIn 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox TextOut 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   4815
   End
   Begin prjNSocks.Server Server 
      Height          =   390
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Client_DataArrival(ByVal Data As String)
TextIn.Text = Data
End Sub

Private Sub Command1_Click()
Client.LocalPort = 1192
Client.RemotePort = 1193
Client.Protocol = sckTCPProtocol
Server.LocalPort = 1193
Server.Protocol = sckTCPProtocol
Client.Connect
End Sub

Private Sub Command2_Click()
Client.Disconnect
End Sub

Private Sub Command3_Click()
Client.SendData TextIn.Text
End Sub

Private Sub Command4_Click()
Server.SendData 1, TextOut.Text
End Sub

Private Sub Form_Load()
MsgBox Client.LocalHostName & "|" & Client.LocalIP
End Sub

Private Sub Server_DataArrival(ByVal Index As Integer, ByVal Data As String)
TextOut.Text = Data
End Sub
