VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Client 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "UserControl1.ctx":000F
   Begin MSWinsockLib.Winsock Sock 
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
      Picture         =   "UserControl1.ctx":0321
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Public vConnectionString As String
Attribute vConnectionString.VB_VarProcData = "ppClient"
'Event Declarations:
Event DataArrival(ByVal Data As String)
Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Event SendComplete()
Event SendProgress(ByVal Percent As Long)
Event Disconnected()

Public Sub About()
frmAbout.Show
frmAbout.InitAboutBox True
End Sub

Private Sub Sock_Close()
RaiseEvent Disconnected
End Sub

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
    Sock.Bind LocalPort, LocalIP
End Sub

Public Sub Disconnect()
    Sock.Close
End Sub

Public Sub Connect()
Attribute Connect.VB_Description = "Connect to the remote computer"
    Sock.Connect
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Sock.GetData a, vbString, bytesTotal
    If a = "Send_Connection_String" Then
    Sock.SendData ConnectionString
    Else
    RaiseEvent DataArrival(a)
    End If
End Sub

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Public Property Get LocalHostName() As String
Attribute LocalHostName.VB_Description = "Returns the local machine name"
    LocalHostName = Sock.LocalHostName
End Property

Public Property Get LocalIP() As String
Attribute LocalIP.VB_Description = "Returns the local machine IP address"
    LocalIP = Sock.LocalIP
End Property

Public Property Get LocalPort() As Long
    LocalPort = Sock.LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    Sock.LocalPort = New_LocalPort
    PropertyChanged "LocalPort"
End Property

Public Property Get ConnectionString() As String
    ConnectionString = vConnectionString
End Property

Public Property Let ConnectionString(ByVal New_ConnectioString As String)
    vConnectionString = New_ConnectioString
    PropertyChanged "ConnectionString"
End Property

Public Property Get Protocol() As ProtocolConstants
Attribute Protocol.VB_Description = "Returns/Sets the socket protocol"
    Protocol = Sock.Protocol
End Property

Public Property Let Protocol(ByVal New_Protocol As ProtocolConstants)
    Sock.Protocol = New_Protocol
    PropertyChanged "Protocol"
End Property

Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
Attribute RemoteHost.VB_ProcData.VB_Invoke_Property = "ppClient"
    RemoteHost = Sock.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    Sock.RemoteHost = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = Sock.RemoteHostIP
End Property

Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
Attribute RemotePort.VB_ProcData.VB_Invoke_Property = "ppClient"
    RemotePort = Sock.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    Sock.RemotePort = New_RemotePort
    PropertyChanged "RemotePort"
End Property

Private Sub Sock_SendComplete()
    RaiseEvent SendComplete
End Sub

Public Sub SendData(ByVal Data As Variant)
Attribute SendData.VB_Description = "Send data to remote computer"
    Sock.SendData Data
End Sub

Private Sub Sock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Percent = (bytesSent / bytesRemaining) * 100
    RaiseEvent SendProgress(Percent)
End Sub

Public Property Get SocketHandle() As Long
Attribute SocketHandle.VB_Description = "Returns the socket handle"
    SocketHandle = Sock.SocketHandle
End Property

Public Property Get State() As Integer
Attribute State.VB_Description = "Returns the state of the socket connection"
    State = Sock.State
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    vConnectionString = PropBag.ReadProperty("ConnectionString", "")
    Sock.LocalPort = PropBag.ReadProperty("LocalPort", 0)
    Sock.Protocol = PropBag.ReadProperty("Protocol", 0)
    Sock.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Sock.RemotePort = PropBag.ReadProperty("RemotePort", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ConnectionString", vConnectionString, "")
    Call PropBag.WriteProperty("LocalPort", Sock.LocalPort, 0)
    Call PropBag.WriteProperty("Protocol", Sock.Protocol, 0)
    Call PropBag.WriteProperty("RemoteHost", Sock.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", Sock.RemotePort, 0)
End Sub

