Attribute VB_Name = "SPECIAL_FOLDERS"
'This program is designed by sherif rofael
'the declarations and function are collected from various _
 places ,
' u can use any part of that program freely ,
' I hope u find the code easy to follow
'mailto: sherif@vbcode.tk
'website: www.vbcode.tk

'Purpose: Access Windows Special Folders

Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type

Private Type ITEMIDLIST
    shellID As SHITEMID
End Type

Const SF_DESKTOP = &H0          'Desktop
Const SF_PROGRAMS = &H2         'Program Files
Const SF_MYDOCS = &H5           'My Documents
Const SF_FAVORITES = &H6        'Favorites; Win 98+
Const SF_STARTUP = &H7
Const SF_RECENT = &H8
Const SF_SENDTO = &H9
Const SF_STARTMENU = &HB
Const SF_MYMUSIC = &HD          'My Music; Win Me+
Const SF_DESKTOP2 = &H10
Const SF_NETHOOD = &H13
Const SF_FONTS = &H14
Const SF_SHELLNEW = &H15
Const SF_STARTUP2 = &H18
Const SF_ALLUSERSDESK = &H19
Const SF_APPDATA = &H1A
Const SF_PRINTHOOD = &H1B
Const SF_APPDATA2 = &H1C
Const SF_TEMPINETFILES = &H20
Const SF_COOKIES = &H21
Const SF_HISTORY = &H22
Const SF_ALLUSERSAPPDATA = &H23
Const SF_WINDOWS = &H24
Const SF_WINSYSTEM = &H25
Const SF_PROGFILES = &H26
Const SF_MYPICS = &H27          'My Pictures: Win Me+
Const SF_USERDIR = &H28
Const SF_WINSYSTEM2 = &H29
Const SF_COMMON = &H2B


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long

Public Function getSpecialFolder(whichFolder As Long, CurrentFormhWnd As Long) As String
    Dim Path As String * 256
    Dim myid As ITEMIDLIST
    Dim rval As Long

    If IsMissing(useForm) Then
    rval = SHGetSpecialFolderLocation(CurrentFormhWnd, whichFolder, myid)
    Else
    rval = SHGetSpecialFolderLocation(CurrentFormhWnd, whichFolder, myid)
    End If
    
    If rval = 0 Then ' If success
      rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal Path)
        If rval Then ' If True
        getSpecialFolder = Left(Path, InStr(Path, Chr(0)) - 1)
        End If
    End If
    
End Function


