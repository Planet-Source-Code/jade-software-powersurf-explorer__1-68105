VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_MenuXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function pGetPicture(sFileName As String) As StdPicture

    '-//=================================================================================
    '-//LOAD THE MENU IMAGES
    '-//=================================================================================
    
    Dim c As Long

    For c = 1 To frmMain.ImgList.ListImages.Count '-// FROM 1 TO Number of Images added into Imagelist
        If LCase(frmMain.ImgList.ListImages(c).Key) = LCase(sFileName) Then
            Set pGetPicture = frmMain.ImgList.ListImages(c).ExtractIcon 'GRABBED IMAGES ICON TO MENU
    'Set pGetPicture = LoadPicture(App.Path + "\Images\" + sFileName + ".ico")

        End If
    Next
    
End Function
