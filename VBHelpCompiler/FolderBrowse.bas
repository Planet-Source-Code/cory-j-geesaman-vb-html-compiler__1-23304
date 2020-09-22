Attribute VB_Name = "FolderBrowse"
'##################################################################
'#I got this module from PSC(http://www.planetsourcecode.com/vb/),#
'#but am sorry to say i forgot where                              #
'##################################################################
Option Explicit
' Maximum long filename path length
Private Const MAX_PATH = 1024

'SendMessage Constants
Private Const BFFM_INITIALIZED = 1
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)

'The Following Constants may be passed to BrowseForFolder
'as vTopFolder or vSelPath
Public Const CSIDL_DESKTOP = &H0    'DeskTop
Public Const CSIDL_PROGRAMS = &H2   'Program Groups Folder
Public Const CSIDL_CONTROLS = &H3   'Control Panel Icons Folder
Public Const CSIDL_PRINTERS = &H4   'Printers Folder
Public Const CSIDL_PERSONAL = &H5   'Documents Folder
Public Const CSIDL_FAVORITES = &H6  'Favorites Folder
Public Const CSIDL_STARTUP = &H7    'Startup Folder
Public Const CSIDL_RECENT = &H8     'Recent folder
Public Const CSIDL_SENDTO = &H9     'SendTo Folder
Public Const CSIDL_BITBUCKET = &HA  'Recycle Bin Folder
Public Const CSIDL_STARTMENU = &HB  'Start Menu Folder
Public Const CSIDL_DESKTOPDIRECTORY = &H10  'Windows\Desktop Folder
Public Const CSIDL_DRIVES = &H11    'Devices Virtual Folder (My Computer)
Public Const CSIDL_NETWORK = &H12   'Network Neighborhood Virtual Folder
Public Const CSIDL_NETHOOD = &H13   'Network Neighborhood Folder
Public Const CSIDL_FONTS = &H14     'Fonts Folder
Public Const CSIDL_TEMPLATES = &H15 'ShellNew folder

Private Type SHItemID
    cb      As Long    'Size of the ID (including cb itself)
    abID    As Byte    'The item ID (variable length)
End Type

Private Type ItemIDList
    mkid    As SHItemID
End Type

Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpCallbackProc  As Long
    lParam          As Long
    iImage          As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ItemIDList) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  Select Case uMsg
    Case BFFM_INITIALIZED
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  End Select
End Function

Public Function BrowseForFolder(hOwnerWnd As Long, Optional ByVal sInstruct As String, Optional vSelPath As Variant, Optional vTopFolder As Variant) As String
' Shows the Browse For Folder dialog
'
' hOwnerWnd     (Long)                     OwnerWindow.hWnd.
' sInstruct     (String)                   Instructions for user.
' vSelPath      (String or CSIDL Constant) Pre-select this Folder.
' vTopFolder    (String or CSIDL Constant) Set the Top folder.
'
' If successful, returns the selected folder's full path,
' returns an empty string otherwise.
'
Dim lRet As Long
Dim pidlRet As Long
Dim sPath As String * MAX_PATH
Dim lItemIDList As ItemIDList
Dim uBrowseInfo As BROWSEINFO
    With uBrowseInfo
        .hOwner = hOwnerWnd
        If IsMissing(vTopFolder) Then
            vTopFolder = CSIDL_DESKTOP
        End If
        If Len(vTopFolder) > 0 And Not IsNumeric(vTopFolder) Then
            .pidlRoot = SHSimpleIDListFromPath(CStr(vTopFolder))
        Else
            lRet = SHGetSpecialFolderLocation(ByVal hOwnerWnd, ByVal CLng(vTopFolder), lItemIDList)
            .pidlRoot = lItemIDList.mkid.cb
        End If
        .lpszTitle = sInstruct
        .lpCallbackProc = FarProc(AddressOf BrowseCallbackProc)
        If IsMissing(vSelPath) Then
            .lParam = .pidlRoot
        ElseIf Len(vSelPath) > 0 And Not IsNumeric(vSelPath) Then
            .lParam = SHSimpleIDListFromPath(CStr(vSelPath))
        Else
            lRet = SHGetSpecialFolderLocation(ByVal hOwnerWnd, ByVal CLng(vSelPath), lItemIDList)
            .lParam = lItemIDList.mkid.cb
        End If
    End With
    pidlRet = SHBrowseForFolder(uBrowseInfo)
    If pidlRet > 0 Then
        If SHGetPathFromIDList(pidlRet, sPath) Then
          BrowseForFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(pidlRet)
    End If
    Call CoTaskMemFree(uBrowseInfo.lParam)
End Function

Public Function FarProc(lpProcName As Long) As Long
    FarProc = lpProcName
End Function
