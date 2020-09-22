VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB->HTML Compiler, By Cory J. Geesaman"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Convert VB To &HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   7080
      Width           =   6975
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Syntax Coloring"
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   6975
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Remove"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   3060
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Edit"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   3360
         Width           =   2535
      End
      Begin VB.CommandButton cmdSyntaxColor 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtSyntaxString 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4200
         TabIndex        =   14
         Top             =   600
         Width           =   2535
      End
      Begin MSComctlLib.ListView lstSyntaxColors 
         Height          =   3255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Color:"
         Height          =   195
         Left            =   4200
         TabIndex        =   20
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&String:"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Colors"
      Height          =   1000
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   6975
      Begin VB.CheckBox Check2 
         Caption         =   "Use BackColor from form or usercontrol the page pertains to if available"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   300
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use BackColor from form or usercontrol the page pertains to if available"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   300
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00400000&
         Height          =   195
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "-When i get source for OLE_Colors->Longs"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   725
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Header Color:"
         Height          =   195
         Left            =   1950
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&BackColor:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Output"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtOutputDir 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Input"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtInputPRJ 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   6135
      End
   End
   Begin MSComDlg.CommonDialog cD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "VB Components(*.vbg, *.vbp, *.frm, *.ctl, *.bas, *.cls, *.pag)|*.vbg;*.vbp;*.frm;*.ctl;*.bas;*.cls;*.pag"
   End
   Begin RichTextLib.RichTextBox tOpener 
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private StrList() As String
Private ColList() As Long
Private LenList() As Long
Private LastV As Long
Private tBStyle As Byte
Public Working As Boolean

Private Sub cmdSyntaxColor_Click()
cD.ShowColor
cmdSyntaxColor.BackColor = cD.Color
txtSyntaxString.ForeColor = cD.Color
End Sub

Private Sub Command1_Click()
Dim OutLctn As String
OutLctn = BrowseForFolder(Me.hWnd, "Browse For HTML Output Directory")
If OutLctn <> "" Then
txtOutputDir.Text = OutLctn
End If
End Sub

Private Sub Command2_Click()
cD.Flags = 0
cD.ShowOpen
If cD.Flags <> 0 Then
txtInputPRJ.Text = cD.FileName
End If
End Sub

Public Sub CheckDir(Dir As String)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Dir) Then
fso.CreateFolder Dir
End If
End Sub

Public Function LineCount(strPath As String) As Long
Dim tStrs As Variant
tOpener.LoadFile strPath
tStrs = Split(tOpener.Text, vbNewLine)
LineCount = UBound(tStrs) + 1
End Function

Public Function ProjectLineCount(strPath As String) As Variant
Dim tS As String, pName As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, b As Variant, hC As String, rPath As String
Dim tForms As Long, tModules As Long, tClassModules As Long, tUserControls As Long, tPropertyPages As Long
Dim lF As Long, lM As Long, lC As Long, lU As Long, lP As Long
a = Split(strPath, "\")
rPath = Mid(strPath, 1, Len(strPath) - Len(a(UBound(a))))
tOpener.LoadFile strPath
rStrs = Split(tOpener.Text, vbNewLine)
tForms = 0
tModules = 0
tClassModules = 0
tUserControls = 0
tPropertyPages = 0
pName = ""
txtOut = ""
i = 0
lF = 0
lM = 0
lC = 0
lU = 0
lP = 0
Do
tS = rStrs(i)
If tS = "" Or tS = " " Then GoTo PastLoop
a = Split(tS, "=")
b = Split(a(1), "; ")
If a(0) = "Name" Then
pName = Mid(a(1), 2, Len(a(1)) - 2)
ElseIf a(0) = "Form" Then
tForms = tForms + 1
lF = lF + LineCount(rPath & a(1))
ElseIf a(0) = "Module" Then
tModules = tModules + 1
lM = lM + LineCount(rPath & b(1))
ElseIf a(0) = "Class" Then
tClassModules = tClassModules + 1
lC = lC + LineCount(rPath & b(1))
ElseIf a(0) = "UserControl" Then
tUserControls = tUserControls + 1
lU = lU + LineCount(rPath & a(1))
ElseIf a(0) = "PropertyPage" Then
tPropertyPages = tPropertyPages + 1
lP = lP + LineCount(rPath & a(1))
End If
PastLoop:
i = i + 1
Loop Until i > UBound(rStrs)
ReDim a(0 To 5)
a(0) = lF + lM + lC + lU + lP
a(1) = tForms
a(2) = tModules
a(3) = tClassModules
a(4) = tUserControls
a(5) = tPropertyPages
ProjectLineCount = a
End Function

Public Sub OutputHTML()
LastV = 0
lstSyntaxColors.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
cmdSyntaxColor.Enabled = False
txtOutputDir.Enabled = False
txtInputPRJ.Enabled = False
Check1.Enabled = False
ReDim StrList(1 To lstSyntaxColors.ListItems.Count) As String
ReDim ColList(1 To lstSyntaxColors.ListItems.Count) As Long
ReDim LenList(1 To lstSyntaxColors.ListItems.Count) As Long
i = 1
Do
StrList(i) = lstSyntaxColors.ListItems(i).Text
ColList(i) = lstSyntaxColors.ListItems(i).ForeColor
LenList(i) = Len(StrList(i))
i = i + 1
Loop Until i > lstSyntaxColors.ListItems.Count
'Open txtInputPRJ.Text For Input As 1
Working = True
Load frmWorking
frmWorking.tBar.Value = 0
frmWorking.pBar.Value = 0
frmWorking.txtFile = "File:"
frmWorking.Show , Me
Select Case LCase(Right(txtInputPRJ.Text, 3))
Case "vbg"
tBStyle = 0
SaveProjectGroup txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "vbp"
tBStyle = 1
SaveProject txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "frm"
tBStyle = 2
SaveForm txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "bas"
tBStyle = 3
SaveModule txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "cls"
tBStyle = 4
SaveClassModule txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "ctl"
tBStyle = 5
SaveUserControl txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
Case "pag"
tBStyle = 6
SavePropertyPage txtInputPRJ.Text, txtOutputDir.Text & "\"
frmWorking.tBar.Value = 0
End Select
Working = False
'Close 1
End Sub

Public Function SaveProjectGroup(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, b As Variant, hC As String, rPath As String
Dim tProjects As Long
Dim P As Long
Dim lP As Long
Dim Projects As Variant
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
rPath = Mid(FileName, 1, Len(FileName) - Len(a(UBound(a))))
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
bgC = LNGtoHEX(Command3.BackColor)
hC = LNGtoHEX(Command4.BackColor)
tProjects = 0
txtOut = ""
i = 0
lP = 0
Do
If Working = False Then Exit Function
tS = rStrs(i)
If tS = "" Or tS = " " Then GoTo PastLoop
a = Split(tS, "=")
If a(0) = "Project" Or a(0) = "StartupProject" Then
tProjects = tProjects + 1
lP = lP + ProjectLineCount(rPath & a(1))(0)
End If
PastLoop:
i = i + 1
If Working = False Then Exit Function
Loop Until i > UBound(rStrs)
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = lP
frmWorking.tBar.Value = 0
ReDim Projects(-1 To tProjects - 1) As Variant
i = 0
Do
If Working = False Then Exit Function
tS = rStrs(i)
If tS = "" Or tS = " " Then GoTo PastLoop2
a = Split(tS, "=")
If a(0) = "Project" Or a(0) = "StartupProject" Then
Projects(P) = SaveProject(rPath & a(1), PathOut)
P = P + 1
End If
PastLoop2:
i = i + 1
If Working = False Then Exit Function
Loop Until i > UBound(rStrs)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & pName & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
If P > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">Projects(" & tProjects & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If Projects(i) <> "" And Projects(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & Projects(i) & ".htm"">" & Projects(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= P
End If
txtOut = txtOut & vbNewLine & "<BR>" & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
a = Split(FileName, "\")
a = a(UBound(a))
a = Split(a, ".")
a = a(0)
If Working = False Then Exit Function
Open PathOut & "index.htm" For Output As 2
Print #2, txtOut
Close 2
SaveProjectGroup = a
End Function

Public Function SaveProject(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, pName As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, b As Variant, hC As String, rPath As String
Dim tForms As Long, tModules As Long, tClassModules As Long, tUserControls As Long, tPropertyPages As Long
Dim F As Long, M As Long, C As Long, U As Long, P As Long
Dim lF As Long, lM As Long, lC As Long, lU As Long, lP As Long
Dim Forms As Variant, Modules As Variant, ClassModules As Variant, UserControls As Variant, PropertyPages As Variant
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
rPath = Mid(FileName, 1, Len(FileName) - Len(a(UBound(a))))
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
bgC = LNGtoHEX(Command3.BackColor)
hC = LNGtoHEX(Command4.BackColor)
tForms = 0
tModules = 0
tClassModules = 0
tUserControls = 0
tPropertyPages = 0
pName = ""
txtOut = ""
i = 0
j = 0
lF = 0
lM = 0
lC = 0
lU = 0
lP = 0
Do
If Working = False Then Exit Function
tS = rStrs(i)
If tS = "" Or tS = " " Then GoTo PastLoop
a = Split(tS, "=")
b = Split(a(1), "; ")
If a(0) = "Name" Then
pName = Mid(a(1), 2, Len(a(1)) - 2)
ElseIf a(0) = "Form" Then
tForms = tForms + 1
lF = lF + LineCount(rPath & a(1))
j = j + 1
ElseIf a(0) = "Module" Then
tModules = tModules + 1
lM = lM + LineCount(rPath & b(1))
j = j + 1
ElseIf a(0) = "Class" Then
tClassModules = tClassModules + 1
lC = lC + LineCount(rPath & b(1))
j = j + 1
ElseIf a(0) = "UserControl" Then
tUserControls = tUserControls + 1
lU = lU + LineCount(rPath & a(1))
j = j + 1
ElseIf a(0) = "PropertyPage" Then
tPropertyPages = tPropertyPages + 1
lP = lP + LineCount(rPath & a(1))
j = j + 1
End If
PastLoop:
i = i + 1
If Working = False Then Exit Function
Loop Until i > UBound(rStrs)
If tBStyle = 1 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = lF + lM + lC + lU + lP
frmWorking.tBar.Value = 0
End If
ReDim Forms(-1 To tForms - 1) As Variant
ReDim Modules(-1 To tModules - 1) As Variant
ReDim ClassModules(-1 To tClassModules - 1) As Variant
ReDim UserControls(-1 To tUserControls - 1) As Variant
ReDim PropertyPages(-1 To tPropertyPages - 1) As Variant
i = 0
Do
If Working = False Then Exit Function
tS = rStrs(i)
If tS = "" Or tS = " " Then GoTo PastLoop2
a = Split(tS, "=")
b = Split(a(1), "; ")
If a(0) = "Form" Then
Forms(F) = SaveForm(rPath & a(1), PathOut)
F = F + 1
ElseIf a(0) = "Module" Then
Modules(M) = SaveModule(rPath & b(1), PathOut)
M = M + 1
ElseIf a(0) = "Class" Then
ClassModules(C) = SaveClassModule(rPath & b(1), PathOut)
C = C + 1
ElseIf a(0) = "UserControl" Then
UserControls(U) = SaveUserControl(rPath & a(1), PathOut)
U = U + 1
ElseIf a(0) = "PropertyPage" Then
PropertyPages(P) = SavePropertyPage(rPath & a(1), PathOut)
P = P + 1
End If
PastLoop2:
i = i + 1
If Working = False Then Exit Function
Loop Until i > UBound(rStrs)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & pName & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
If F > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">Forms(" & tForms & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If Forms(i) <> "" And Forms(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & Forms(i) & ".htm"">" & Forms(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= F
End If
If M > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">Modules(" & tModules & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If Modules(i) <> "" And Modules(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & Modules(i) & ".htm"">" & Modules(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= M
End If
If C > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">Class Modules(" & tClassModules & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If ClassModules(i) <> "" And ClassModules(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & ClassModules(i) & ".htm"">" & ClassModules(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= C
End If
If U > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">User Controls(" & tUserControls & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If UserControls(i) <> "" And UserControls(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & UserControls(i) & ".htm"">" & UserControls(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= U
End If
If P > 0 Then
txtOut = txtOut & vbNewLine & "<BR>" & "<FONT SIZE=""4"" COLOR=""" & hC & _
""">Property Pages(" & tPropertyPages & ")</FONT>" & vbNewLine & "<HR SIZE=""1"" WIDTH=""100%"">"
i = 0
Do
If Working = False Then Exit Function
If PropertyPages(i) <> "" And PropertyPages(i) <> " " Then txtOut = txtOut & vbNewLine & "<A HREF=""" & PropertyPages(i) & ".htm"">" & PropertyPages(i) & "</A>" & vbNewLine & "<BR>"
i = i + 1
If Working = False Then Exit Function
Loop Until i >= P
End If
txtOut = txtOut & vbNewLine & "<BR>" & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & pName & ".htm" For Output As 2
Print #2, txtOut
Close 2
SaveProject = pName
End Function

Public Function SaveForm(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, hC As String
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
j = 0
i = 0
HitAttrib = False
frmWorking.pBar.Min = 0
frmWorking.pBar.Max = UBound(rStrs)
frmWorking.pBar.Value = 0
If tBStyle = 2 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = UBound(rStrs)
frmWorking.tBar.Value = 0
End If
If Check1.Value = vbChecked Then
bgC = LNGtoHEX(&HE0E0E0)
Else
bgC = LNGtoHEX(Command3.BackColor)
End If
hC = LNGtoHEX(Command4.BackColor)
Do
tS = rStrs(i + j)
'Line Input #1, tS
'''''start bg bs
If Check1.Value = vbChecked Then
If LCase(Mid(tS, 1, 14)) = "begin vb.form " Then
OfD = True
End If
If (InStr(1, LCase(tS), "end", vbTextCompare) > 0) Then
OfD = False
End If
If (InStr(1, LCase(tS), "backcolor", vbTextCompare) > 0) And OfD = True Then
a = Split(tS, "=")
b = NoSpace(a(1))
bgC = "#" & Mid(b, 4)
If Len(bgC) < 7 Then
Do
bgC = bgC & "0"
Loop Until Len(bgC) > 6
End If
bgC = "#" & Mid(bgC, 6, 2) & Mid(bgC, 4, 2) & Mid(bgC, 2, 2)
End If
'''''end bg bs
End If
If Mid(tS, 1, 13) = "Attribute VB_" Then
HitAttrib = True
If Mid(tS, 1, 18) = "Attribute VB_Name " Then
fo = Mid(tS, 22, Len(tS) - 22)
End If
j = j + i
i = 0
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & fo & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
GoTo After_Loop
End If
If HitAttrib = True Then
If Working = False Then Exit Function
txtOut = txtOut & vbNewLine & "<br>" & vbNewLine & FormattedLine(tS)
If Working = False Then Exit Function
End If
After_Loop:
frmWorking.pBar.Value = i + j
If UpdateTBar = False Then Exit Function
i = i + 1
Loop Until i + j > UBound(rStrs) ' EOF(1)
frmWorking.pBar.Value = 0
txtOut = txtOut & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & fo & ".htm" For Output As 2
Print #2, txtOut
Close 2
SaveForm = fo
End Function

Public Function SaveModule(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, hC As String
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
j = 0
i = 0
HitAttrib = False
frmWorking.pBar.Min = 0
frmWorking.pBar.Max = UBound(rStrs)
frmWorking.pBar.Value = 0
If tBStyle = 3 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = UBound(rStrs)
frmWorking.tBar.Value = 0
End If
bgC = LNGtoHEX(Command3.BackColor)
Do
tS = rStrs(i + j)
'Line Input #1, tS
If Mid(tS, 1, 13) = "Attribute VB_" Then
HitAttrib = True
If Mid(tS, 1, 18) = "Attribute VB_Name " Then
fo = Mid(tS, 22, Len(tS) - 22)
End If
j = j + i
i = 0
hC = LNGtoHEX(Command4.BackColor)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & fo & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
GoTo After_Loop
End If
If HitAttrib = True Then
If Working = False Then Exit Function
txtOut = txtOut & vbNewLine & "<br>" & vbNewLine & FormattedLine(tS)
If Working = False Then Exit Function
End If
After_Loop:
frmWorking.pBar.Value = i + j
If UpdateTBar = False Then Exit Function
i = i + 1
Loop Until i + j > UBound(rStrs) ' EOF(1)
frmWorking.pBar.Value = 0
txtOut = txtOut & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & fo & ".htm" For Output As 2
Print #2, txtOut
Close 2
SaveModule = fo
End Function

Public Function SaveClassModule(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, hC As String
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
j = 0
i = 0
HitAttrib = False
frmWorking.pBar.Min = 0
frmWorking.pBar.Max = UBound(rStrs)
frmWorking.pBar.Value = 0
If tBStyle = 4 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = UBound(rStrs)
frmWorking.tBar.Value = 0
End If
bgC = LNGtoHEX(Command3.BackColor)
Do
tS = rStrs(i + j)
'Line Input #1, tS
If Mid(tS, 1, 13) = "Attribute VB_" Then
HitAttrib = True
If Mid(tS, 1, 18) = "Attribute VB_Name " Then
fo = Mid(tS, 22, Len(tS) - 22)
End If
j = j + i
i = 0
hC = LNGtoHEX(Command4.BackColor)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & fo & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
GoTo After_Loop
End If
If HitAttrib = True Then
If Working = False Then Exit Function
txtOut = txtOut & vbNewLine & "<br>" & vbNewLine & FormattedLine(tS)
If Working = False Then Exit Function
End If
After_Loop:
frmWorking.pBar.Value = i + j
If UpdateTBar = False Then Exit Function
i = i + 1
Loop Until i + j > UBound(rStrs) ' EOF(1)
frmWorking.pBar.Value = 0
txtOut = txtOut & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & fo & ".htm" For Output As 2
Print #2, txtOut
Close 2
SaveClassModule = fo
End Function

Public Function SaveUserControl(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, hC As String
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
j = 0
i = 0
HitAttrib = False
frmWorking.pBar.Min = 0
frmWorking.pBar.Max = UBound(rStrs)
frmWorking.pBar.Value = 0
If tBStyle = 5 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = UBound(rStrs)
frmWorking.tBar.Value = 0
End If
If Check1.Value = vbChecked Then
bgC = LNGtoHEX(&HE0E0E0)
Else
bgC = LNGtoHEX(Command3.BackColor)
End If
Do
tS = rStrs(i + j)
'Line Input #1, tS
'''''start bg bs
If Check1.Value = vbChecked Then
If LCase(Mid(tS, 1, 14)) = "begin vb.usercontrol " Then
OfD = True
End If
If (InStr(1, LCase(tS), "end", vbTextCompare) > 0) Then
OfD = False
End If
If (InStr(1, LCase(tS), "backcolor", vbTextCompare) > 0) And OfD = True Then
a = Split(tS, "=")
b = NoSpace(a(1))
bgC = "#" & Mid(b, 4)
If Len(bgC) < 7 Then
Do
bgC = bgC & "0"
Loop Until Len(bgC) > 6
End If
bgC = "#" & Mid(bgC, 6, 2) & Mid(bgC, 4, 2) & Mid(bgC, 2, 2)
End If
'''''end bg bs
End If
If Mid(tS, 1, 13) = "Attribute VB_" Then
HitAttrib = True
If Mid(tS, 1, 18) = "Attribute VB_Name " Then
fo = Mid(tS, 22, Len(tS) - 22)
End If
j = j + i
i = 0
hC = LNGtoHEX(Command4.BackColor)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & fo & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
GoTo After_Loop
End If
If HitAttrib = True Then
If Working = False Then Exit Function
txtOut = txtOut & vbNewLine & "<br>" & vbNewLine & FormattedLine(tS)
If Working = False Then Exit Function
End If
After_Loop:
frmWorking.pBar.Value = i + j
If UpdateTBar = False Then Exit Function
i = i + 1
Loop Until i + j > UBound(rStrs) ' EOF(1)
frmWorking.pBar.Value = 0
txtOut = txtOut & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & fo & ".htm" For Output As 2
Print #2, txtOut
Close 2
SaveUserControl = fo
End Function

Public Function SavePropertyPage(FileName As String, PathOut As String) As String
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
Exit Function
AfterEF:
Dim tS As String, txtOut As String, HitAttrib As Boolean, rStrs As Variant, OfD As Boolean, bgC As String, a As Variant, hC As String
a = Split(FileName, "\")
frmWorking.txtFile = a(UBound(a)) & ":"
tOpener.LoadFile FileName
rStrs = Split(tOpener.Text, vbNewLine)
j = 0
i = 0
HitAttrib = False
frmWorking.pBar.Min = 0
frmWorking.pBar.Max = UBound(rStrs)
frmWorking.pBar.Value = 0
If tBStyle = 6 Then
frmWorking.tBar.Min = 0
frmWorking.tBar.Max = UBound(rStrs)
frmWorking.tBar.Value = 0
End If
If Check1.Value = vbChecked Then
bgC = LNGtoHEX(&HE0E0E0)
Else
bgC = LNGtoHEX(Command3.BackColor)
End If
Do
tS = rStrs(i + j)
'Line Input #1, tS
'''''start bg bs
If Check1.Value = vbChecked Then
If LCase(Mid(tS, 1, 21)) = "begin vb.propertypage " Then
OfD = True
End If
If (InStr(1, LCase(tS), "end", vbTextCompare) > 0) Then
OfD = False
End If
If (InStr(1, LCase(tS), "backcolor", vbTextCompare) > 0) And OfD = True Then
a = Split(tS, "=")
b = NoSpace(a(1))
bgC = "#" & Mid(b, 4)
If Len(bgC) < 7 Then
Do
bgC = bgC & "0"
Loop Until Len(bgC) > 6
End If
bgC = "#" & Mid(bgC, 6, 2) & Mid(bgC, 4, 2) & Mid(bgC, 2, 2)
End If
'''''end bg bs
End If
If Mid(tS, 1, 13) = "Attribute VB_" Then
HitAttrib = True
If Mid(tS, 1, 18) = "Attribute VB_Name " Then
fo = Mid(tS, 22, Len(tS) - 22)
End If
j = j + i
i = 0
hC = LNGtoHEX(Command4.BackColor)
txtOut = "<HTML>" & vbNewLine & "<BODY BGCOLOR=""" & bgC & """ TEXT=""" & LNGtoHEX(lstSyntaxColors.ListItems.Item("ForCol").ForeColor) & """>" _
& vbNewLine & "<FONT SIZE=""5"" COLOR=""" & hC & """>" & fo & "</FONT>" & vbNewLine & "<HR SIZE=""10"" WIDTH=""100%"">"
GoTo After_Loop
End If
If HitAttrib = True Then
If Working = False Then Exit Function
txtOut = txtOut & vbNewLine & "<br>" & vbNewLine & FormattedLine(tS)
If Working = False Then Exit Function
End If
After_Loop:
frmWorking.pBar.Value = i + j
If UpdateTBar = False Then Exit Function
i = i + 1
Loop Until i + j > UBound(rStrs) ' EOF(1)
frmWorking.pBar.Value = 0
txtOut = txtOut & vbNewLine & "<HR SIZE=""10"" WISTH=""100%"">This page was created by the VB->HTML Compiler, made by Cory J. Geesaman, <A HREF=""http://www.naven.net/"">http://www.naven.net/</A>" & "</BODY>" & vbNewLine & "</HTML>"
CheckDir PathOut
If Working = False Then Exit Function
Open PathOut & fo & ".htm" For Output As 2
Print #2, txtOut
Close 2
SavePropertyPage = fo
End Function

Public Function NoSpace(ByVal TextIn)
Dim tS As String
i = 1
Do
If Mid(TextIn, i, 1) <> " " And Mid(TextIn, i, 1) <> "&" Then tS = tS & Mid(TextIn, i, 1)
i = i + 1
Loop Until i > Len(TextIn)
NoSpace = tS
End Function

Public Function FormattedLine(TextIn As String) As String
Dim tS As String, tL As Long, tmpL As Long, tmpS As String, InString As Boolean, oM As Boolean
oM = False
'TextIn = Replace(TextIn, "&", "&amp;")This was causing more probs than it fixed
TextIn = Replace(TextIn, "<", "&lt")
TextIn = Replace(TextIn, ">", "&gt")
'TextIn = Replace(TextIn, " ", "&nbsp;")This was causing more probs than it fixed
tL = Len(TextIn)
InString = False
i = 1
Do
If Mid(TextIn, i, 1) = """" Then InString = Not InString
If InString = False Then '0
If Mid(TextIn, i, 1) = "'" Then '1
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(lstSyntaxColors.ListItems("Comment").ForeColor) & """>" & Mid(TextIn, i) & "</font>"
GoTo FuncDone
End If '1
j = 1
Do
tmpS = StrList(j)
tmpC = ColList(j)
tmpL = LenList(j)
If tmpS = "" Or tmpS = " " Then GoTo NextJ

If (tL = tmpL) Or (i <= 1 And tL >= i + tmpL + 1) Or (i > 1 And tL >= i + tmpL) Then '1
On Error GoTo GoIn1
If tL = tmpL Then GoTo GoIn1
If (TextIn = tmpS) Or _
(i = tL - tmpL And (Mid(TextIn, i, tmpL + 1) = ";" & tmpS Or Mid(TextIn, i, tmpL + 1) = "." & tmpS Or Mid(TextIn, i, tmpL + 1) = "(" & tmpS)) Or _
(i <= 1 And (Mid(TextIn, i, tmpL + 1) = tmpS & " " Or Mid(TextIn, i, tmpL + 1) = tmpS & "," Or Mid(TextIn, i, tmpL + 1) = tmpS & ")")) Or _
(i > 1 And (Mid(TextIn, i, tmpL + 2) = " " & tmpS & " " Or Mid(TextIn, i, tmpL + 2) = " " & tmpS & "," Or Mid(TextIn, i, tmpL + 2) = " " & tmpS & ")" Or Mid(TextIn, i, tmpL + 2) = "." & tmpS & " " Or Mid(TextIn, i, tmpL + 2) = "." & tmpS & "," Or Mid(TextIn, i, tmpL + 2) = "." & tmpS & ")") Or Mid(TextIn, i, tmpL + 2) = "(" & tmpS & " " Or Mid(TextIn, i, tmpL + 2) = "(" & tmpS & "," Or Mid(TextIn, i, tmpL + 2) = "(" & tmpS & ")") Then '2
GoIn1:
On Error GoTo ErrorH
If TextIn = tmpS Then
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(tmpC) & """>" & TextIn & "</font>"
i = i + tmpL
GoTo PastCrap
Else
If i <= 1 Then
If i = tL - tmpL - 1 Then
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(tmpC) & """>" & Mid(TextIn, tL - tmpL - 1, tmpL + 1) & "</font>"
i = i + tmpL + 1
GoTo PastCrap
Else
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(tmpC) & """>" & Mid(TextIn, i, tmpL) & "</font>"
i = i + tmpL
GoTo PastCrap
End If
Else
If i = tL - tmpL Then
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(tmpC) & """>" & Mid(TextIn, tL - tmpL, tmpL + 2) & "</font>"
i = i + tmpL + 1
GoTo PastCrap
Else
tS = tS & "<FONT NAME=""Courier New"" COLOR=""" & LNGtoHEX(tmpC) & """>" & Mid(TextIn, i, tmpL + 2) & "</font>"
i = i + tmpL + 1
GoTo PastCrap2
End If
End If
End If
End If '2
End If '1
NextJ:
j = j + 1
Loop Until j > UBound(StrList)
End If '0
If oM = False Then tS = tS & Mid(TextIn, i, 1)
i = i + 1
GoTo PastCrap
PastCrap2:
oM = True
GoTo PastCrap3
PastCrap:
oM = False
PastCrap3:
Loop Until i > Len(TextIn)
FuncDone:
FormattedLine = tS
Exit Function
ErrorH:
Call Error
End Function

'##############use this if u do not have the replace function########
'Public Function Replace(tS As String,sFind as string,sReplace as string) As String
'Dim tOut As String
'i = 1
'Do
'If tOut - i >= Len(sfind) Then
'If Mid(tS, i, len(sfind)) = sFind Then
'tOut = tOut & sreplace
'Else
'tOut = tOut & Mid(tS, i, 1)
'End If
'Else
'tout=tout & mid(ts,i,1)
'End If
'i = i + 1
'Loop Until i > Len(tS)
'Replace = tOut
'End Function
'#######################################################################

Private Sub Command3_Click()
cD.ShowColor
Command3.BackColor = cD.Color
lstSyntaxColors.BackColor = cD.Color
End Sub

Private Sub Command4_Click()
cD.ShowColor
Command4.BackColor = cD.Color
End Sub

Private Sub Command5_Click()
Dim a As ListItem
If Not lstSyntaxColors.SelectedItem Is Nothing Then
lstSyntaxColors.SelectedItem.Text = txtSyntaxString.Text
lstSyntaxColors.SelectedItem.ForeColor = txtSyntaxString.ForeColor
End If
End Sub

Private Sub Command6_Click()
If Not lstSyntaxColors.SelectedItem Is Nothing Then
Command7.Default = True
lstSyntaxColors.ListItems.Remove lstSyntaxColors.SelectedItem.Index
End If
End Sub

Private Sub Command7_Click()
Dim tLI As ListItem
If txtSyntaxString.Text = "" Or txtSyntaxString.Text = " " Then Exit Sub
If txtSyntaxString = "'Normal Color" Then
Set tLI = lstSyntaxColors.ListItems.Add(, "ForCol", txtSyntaxString.Text)
ElseIf txtSyntaxString = "'Comment" Then
Set tLI = lstSyntaxColors.ListItems.Add(, "Comment", txtSyntaxString.Text)
Else
Set tLI = lstSyntaxColors.ListItems.Add(, , txtSyntaxString.Text)
End If
tLI.ForeColor = txtSyntaxString.ForeColor
txtSyntaxString.Text = ""
End Sub

Private Sub Command8_Click()
OutputHTML
End Sub

Private Sub Form_Load()
Dim tLI As ListItem
Set tLI = lstSyntaxColors.ListItems.Add(, , "Get")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Let")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Set")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "If")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "ElseIf")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Else")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "For")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "To")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Do")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Loop")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "While")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Wend")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Until")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Sub")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Function")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Property")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Private")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Public")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "ByVal")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "As")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Then")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Dim")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "New")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "ReDim")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Preserve")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "True")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "False")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Is")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Not")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Nothing")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "End")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Exit")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "String")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Integer")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Long")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Byte")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Double")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Single")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Varient")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Boolean")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CInt")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CStr")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CLng")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CByte")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CVar")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CDbl")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "CSng")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Input")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Open")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Close")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Select")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Case")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Binary")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Put")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Output")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Print")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "B")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "And")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Or")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Call")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Event")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "RaiseEvent")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "On")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Error")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "GoTo")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Line")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Circle")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "PSet")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Scale")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Option")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Explicit")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Declare")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Lib")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "UBound")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "LBound")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Type")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, , "Enum")
tLI.ForeColor = &H800000
Set tLI = lstSyntaxColors.ListItems.Add(, "Comment", "'Comment")
tLI.ForeColor = &H8000&
Set tLI = lstSyntaxColors.ListItems.Add(, "ForCol", "'Normal Color")
tLI.ForeColor = 0
Form_Resize
End Sub

Private Sub Form_Resize()
lstSyntaxColors.ColumnHeaders(1).Width = lstSyntaxColors.Width - 260
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstSyntaxColors_ItemClick(ByVal Item As MSComctlLib.ListItem)
Command5.Default = True
txtSyntaxString.Text = Item.Text
txtSyntaxString.ForeColor = Item.ForeColor
cmdSyntaxColor.BackColor = Item.ForeColor
End Sub

Public Function LNGtoHEX(ByVal lColor As Long, Optional bProper As Boolean = True) As String
    Dim i As Integer
    Dim bytRGB(2) As Byte
    Dim sHex As String
    CopyMemory bytRGB(0), lColor, 3
    For i = 0 To 2
        sHex = sHex & IIf(bytRGB(i) < 16, "0", "") & Hex$(bytRGB(i))
    Next
    If bProper Then sHex = "#" & sHex
    LNGtoHEX = sHex
End Function

Private Function UpdateTBar() As Boolean
On Error GoTo ExitFunc
GoTo AfterEF
ExitFunc:
UpdateTBar = False
Exit Function
AfterEF:
DoEvents
If tBStyle > 1 Then
If frmWorking.pBar.Value <> LastV Then
If LastV < frmWorking.pBar.Value Then
frmWorking.tBar.Value = frmWorking.tBar.Value + (frmWorking.pBar.Value - LastV)
Else
frmWorking.tBar.Value = frmWorking.tBar.Value + LastV
End If
LastV = frmWorking.pBar.Value
End If
UpdateTBar = True
Else
frmWorking.tBar.Value = frmWorking.tBar.Value + 1
UpdateTBar = True
End If

End Function

