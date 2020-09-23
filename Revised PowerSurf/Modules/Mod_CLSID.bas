Attribute VB_Name = "Mod_CLSID"
Option Explicit
'-//=============================================================================================================================
'-// API FOR GETTING SPECIAL FOLDER
'-//=============================================================================================================================


    '-//=============================================================================================================================
    '-// CONSTANT VALUE (FROM API GUIDE)
    '-//=============================================================================================================================
    Const CLSID_STARTMENU = &HB
    Const CLSID_DESKTOP = &H0
    Const CLSID_RECENT = &H8
    Const CLSID_PROGRAMS = &H2
    Const CLSID_FAVORITES = &H6
    Const CLSID_STARTUP = &H7
    Const CLSID_NETHOOD = &H13
    '/+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '| //- END -//
    '\+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

 Function GetSpecialfolder(CLSID As Long) As String
    Dim r, NOERROR As Long
    Dim IDL As ITEMIDLIST
    Dim Path$
    
    '-//=====================================================================================
    '-// GET THE SPECIAL FOLDER
    '-//=====================================================================================
    r = SHGetSpecialFolderLocation(100, CLSID, IDL)
    If r = NOERROR Then
          '-//===============================================================================
          '-// Create a buffer
          '-//===============================================================================
        Path$ = Space$(512)
        
          '-//================================================================================
          '-// Get the path from the IDList
          '-//================================================================================
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
          
        '-//==================================================================================
        '-// Remove the unnecessary chr$(0)'s
        '-//==================================================================================
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Sub NavStartMenu()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_STARTMENU)
End Sub

Sub NavDesktop()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_DESKTOP)
End Sub

Sub NavRecent()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_RECENT)
End Sub

Sub NavPrograms()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_PROGRAMS)
End Sub

Sub NavFavorites()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_FAVORITES)
End Sub

Sub NavStartUp()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_STARTUP)
End Sub

Sub NavNetHood()
    frmMain.WB.Navigate GetSpecialfolder(CLSID_NETHOOD)
End Sub


