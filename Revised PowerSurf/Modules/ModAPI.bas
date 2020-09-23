Attribute VB_Name = "ModAPI"
'-// API DECLARATIONS (Application Programming Interface)
'-//=====================================================================================
'-// API FOR OPENING WEBSITE, FILE  AND OTHERS
'-//=====================================================================================
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'-//=====================================================================================
'-// API TO DOWNLOAD FILE
'-//=====================================================================================
Public Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
