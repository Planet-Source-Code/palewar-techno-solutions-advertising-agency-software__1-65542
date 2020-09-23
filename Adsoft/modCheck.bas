Attribute VB_Name = "modCheck"
Option Explicit
Private Const S_OK = &H0                ' Success
Private Const S_FALSE = &H1             ' The Folder is valid, but does not exist
Private Const E_INVALIDARG = &H80070057 ' Invalid CSIDL Value

Private Const CSIDL_LOCAL_APPDATA = &H1C&
Private Const CSIDL_FLAG_CREATE = &H8000&

Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1
Private Const MAX_PATH = 260

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Enum Folders
 Desktop = &H0
 Internet = &H1
 Programs = &H2
 ControlsFolder = &H3
 Printers = &H4
 Personal = &H5
 Favorites = &H6
 StartUp = &H7
 Recent = &H8
 SendTo = &H9
 RecycleBin = &HA
 StartMenu = &HB
 DesktopDirectory = &H10
 Drives = &H11
 Network = &H12
 Nethood = &H13
 Fonts = &H14
 Templates = &H15
 Common_StartMenu = &H16
 Common_Programs = &H17
 Common_StartUp = &H18
 Common_DesktopDirectory = &H19
 ApplicationData = &H1A
 PrintHood = &H1B
 AltStartUp = &H1D
 Common_AltStartUp = &H1E
 Common_Favorites = &H1F
 InternetCache = &H20
 Cookies = &H21
 History = &H22
End Enum

'Check special folder locations through API, returns their path if they exist
Function CheckFolderID(Folder As Folders) As String
Dim sPath As String
Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH, 0)

RetVal = SHGetFolderPath(0, Folder Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)

Select Case RetVal
    Case S_OK
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        CheckFolderID = sPath
    Case S_FALSE
        CheckFolderID = ""
    Case E_INVALIDARG
         CheckFolderID = ""
End Select
End Function
