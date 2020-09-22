VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjNSocks.Server Server1 
      Height          =   390
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin prjNSocks.Client Client1 
      Height          =   390
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      ConnectionString=   "gbvhm,km"
      LocalPort       =   456
      RemoteHost      =   "54667"
      RemotePort      =   245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
