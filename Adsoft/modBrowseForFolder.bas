Attribute VB_Name = "modBrowseForFolder"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Enum BrowseForFolderFlags
    ReturnFileSystemFoldersOnly = &H1
    DontGoBelowDomain = &H2
    IncludeStatusText = &H4
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    IncludeTextBox = &H10
    ReturnFileSystemAncestors = &H8
End Enum

Private Type BrowseInfo
     hwndOwner As Long
     pidlroot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Dim pidlroot As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (hwnd As Long, nFolder As Long, hToken As Long, dwReserved As Long, ppidl As Long) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Function BrowseForFolder(hwnd As Long, Optional Title As String, Optional Flags As BrowseForFolderFlags, Optional StartUpSpecialFolder As Folders) As String

 
     Dim iNull As Integer
     Dim IDList As Long
     Dim Result As Long
     Dim Path As String
     Dim bi As BrowseInfo
     Dim Ret As String
     If Flags = 0 Then Flags = BIF_RETURNONLYFSDIRS
     
    'Type Settings
     With bi
        Ret = CheckFolderID(StartUpSpecialFolder) 'Check if the special folder exists
        If Ret <> "" Then .pidlroot = StartUpSpecialFolder 'If there is any valid ID use it
        .hwndOwner = hwndOwner 'Set Owner Window
        .ulFlags = Flags 'Set flags (if any)
        .lpszTitle = lstrcat(Title, Chr(0)) 'Append title string to a long value
     End With

    'Execute the BrowseForFolder shell API and display the dialog
     IDList = SHBrowseForFolder(bi) 'Return ID List (selected path in a long value)
     
    'Get the info out of the dialog
     If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path) 'Convert ID list to a string
        iNull = InStr(Path, vbNullChar) 'Get the position of the null character
        If iNull Then Path = Left$(Path, iNull - 1) 'Remove the null character
     End If

    'If Cancel button was clicked, error occured or Non File System Folder was selected then Path = ""
     BrowseForFolder = Path
End Function
