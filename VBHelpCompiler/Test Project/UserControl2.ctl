VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Server 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "UserControl2.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "UserControl2.ctx":000F
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   2160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   0
      Picture         =   "UserControl2.ctx":0321
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event DataArrival(ByVal Index As Integer, ByVal Data As String)
Event SendComplete()
Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Event Disconnected(ByVal Index As Integer)
Event Connected(ByVal Index As Integer)
Private vCurrentUsers As Integer, vMaxUsers As Integer

Private Function GetFreeSock() As Integer

End Function

Public Sub About()
Attribute About.VB_UserMemId = -552
frmAbout.Show
frmAbout.InitAboutBox False
End Sub

Private Sub Sock_Close(Index As Integer)
RaiseEvent Disconnected(Index)
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
a = GetFreeSock
Sock(a).Accept requestID
Sock(a).SendData "Send_Connection_String"
RaiseEvent Connected(a)
End Sub

Public Property Get Protocol() As ProtocolConstants
    Protocol = Sock(0).Protocol
End Property

Public Property Let Protocol(ByVal New_Protocol As ProtocolConstants)
    Sock(0).Protocol = New_Protocol
    PropertyChanged "Protocol"
End Property

Private Sub UserControl_Initialize()
UserControl_Resize
End Sub

Private Sub UserControl_Paint()
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
If UserControl.Width <> 385 Then UserControl.Width = 385
If UserControl.Height <> 385 Then UserControl.Height = 385
End Sub

Public Sub Bind(Optional ByVal LocalPort As Variant, Optional ByVal LocalIP As Variant)
Attribute Bind.VB_Description = "Binds socket to specific port and adapter"
    Sock(0).Bind LocalPort, LocalIP
End Sub

Public Property Get BytesReceived() As Long
Attribute BytesReceived.VB_Description = "Returns the number of bytes received on this connection"
    BytesReceived = Sock(0).BytesReceived
End Property

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Sock(Index).GetData Data, vbString, bytesTotal
    RaiseEvent DataArrival(Index, Data)
End Sub

Public Sub Listen()
Attribute Listen.VB_Description = "Listen for incoming connection requests"
    Sock(0).Listen
End Sub

Public Property Get LocalHostName() As String
Attribute LocalHostName.VB_Description = "Returns the local machine name"
    LocalHostName = Sock(0).LocalHostName
End Property

Public Property Get LocalIP() As String
Attribute LocalIP.VB_Description = "Returns the local machine IP address"
    LocalIP = Sock(0).LocalIP
End Property

Public Property Get LocalPort() As Long
Attribute LocalPort.VB_Description = "Returns/Sets the port used on the local computer"
    LocalPort = Sock(0).LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    Sock(0).LocalPort() = New_LocalPort
    PropertyChanged "LocalPort"
End Property

Public Property Get RemoteHost(Index As Integer) As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = Sock(Index).RemoteHost
End Property

Public Property Get RemoteHostIP(Index As Integer) As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = Sock(Index).RemoteHostIP
End Property

Private Sub Sock_SendComplete(Index As Integer)
    RaiseEvent SendComplete
End Sub

Public Sub SendData(Index As Integer, ByVal Data As String)
Attribute SendData.VB_Description = "Send data to remote computer"
    Sock(0).SendData Data
End Sub

Private Sub Sock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    RaiseEvent SendProgress(Index, bytesSent, bytesRemaining)
End Sub

Public Property Get SocketHandle(Index As Integer) As Long
Attribute SocketHandle.VB_Description = "Returns the socket handle"
    SocketHandle = Sock(Index).SocketHandle
End Property

Public Property Get State(Index As Integer) As Integer
Attribute State.VB_Description = "Returns the state of the socket connection"
    State = Sock(Index).State
End Property

Public Sub Disconnect(Index As Integer)
Attribute Disconnect.VB_Description = "Close current connection"
    Sock(Index).Close
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Sock(0).LocalPort = PropBag.ReadProperty("LocalPort", 0)
    Sock(0).RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Sock(0).Protocol = PropBag.ReadProperty("Protocol", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("LocalPort", Sock(0).LocalPort, 0)
    Call PropBag.WriteProperty("Protocol", Sock(0).Protocol, 0)
    Call PropBag.WriteProperty("RemoteHost", Sock(0).RemoteHost, "")
End Sub
