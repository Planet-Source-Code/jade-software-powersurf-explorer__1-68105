Attribute VB_Name = "GlobalMod"
Option Explicit

'-//====================================================================================================================================================
'-// VARIABLE FOR ADO CONNECTION
'-//====================================================================================================================================================
Global CN                As New ADODB.Connection   '-// FOR CREATE NEW ADO CONNECTION

'-//====================================================================================================================================================
'-// FOR PROFILE INFORMATION
'-//====================================================================================================================================================
Public RsProfile         As New ADODB.Recordset    '-// FOR PROFILE INFORMATION
Public RsloadPro         As New ADODB.Recordset    '-// FOR LOADING COUNTRY INFORMATION
Public hist              As Boolean                '-// FOR HISTORY VIEWING
Public sPlay             As Boolean                '-// FOR WINMEDIA
Public rsAutoComplete    As New ADODB.Recordset    '-// FOR AUTOCOMPLETE COMBO BOX
Public rsAdditem         As New ADODB.Recordset    '-// FOR ADDING URL ON CBOADDRESS
Public rsReloadURL       As New ADODB.Recordset    '-// FOR RELOAD AUTOCOMPLETE

'-//====================================================================================================================================================
'-// FOR LOCK APPLICATION
'-//====================================================================================================================================================
Public rschkPass         As New ADODB.Recordset     '-// FOR CHECKING PASSWORD IF EXIST
Public rsAddPass         As New ADODB.Recordset     '-// FOR SAVING PASSWORD FOR LOCK
Public rsLock            As New ADODB.Recordset     '-// FOR LOCKING APPLICATION
Public rschangePass      As New ADODB.Recordset     '-// FOR CHANGE LOCK PASSWORD

Public rsBlock           As New ADODB.Recordset     '-// FOR BANNED URL
Public rsBlocksite       As New ADODB.Recordset     '-// FOR ADDING URL TO BANNED SITE
Public RSDelBlock        As New ADODB.Recordset     '-// FOR DELETING BLOCKED WEBSITE
Public RSLoadBlock       As New ADODB.Recordset     '-// FOR LOADING BLOCKED WEBSITE

Public RsBLockPass       As New ADODB.Recordset     '-// VERIFY PASSWORD FOR ADDING URL TO BANNED

'-//====================================================================================================================================================
'-// CREATE DATABASE CONNECTION
'-//====================================================================================================================================================
Sub get_connected()
Dim gs As String
On Error GoTo suRedb

   CN.Open "Provider=Microsoft.jet.oledb.4.0;data source=" _
   & App.Path & "/DB/DBURL.dat; jet oledb:database password=246"
  
   Exit Sub
   
suRedb:
   gs = "Either Database does not exist or" + vbCrLf
   gs = gs + "Database password has changed."
   MsgBox gs, vbCritical
   End
   
End Sub

'-//====================================================================================================================================================
'-// FOR VIEWING HISTORY
'-//====================================================================================================================================================
 Sub ViewHistory()

If hist = True Then

        hist = False
        frmMain.WBFav.Visible = True
        frmMain.CoolBarFav.Visible = True
        frmMain.cmdClose.Visible = True
        
        frmMain.WBFav.Left = 80
        frmMain.CoolBarFav.Left = 80
        frmMain.WBFav.Top = 2185
        frmMain.WBFav.Height = frmMain.ScaleHeight - 2530
            
        frmMain.WB.Left = 2555
        frmMain.WB.Height = frmMain.ScaleHeight - 2145
        frmMain.WB.Width = 12735
        frmMain.WBFav.Navigate "C:\Windows\History"
   
        
End If
    
End Sub

Sub viewplayer()

If sPlay = True Then
    
    sPlay = False
        
        frmMain.Picture1.Visible = True
        frmMain.Picture1.Left = 80
        frmMain.Picture1.Top = 1820
        frmMain.Picture1.Height = frmMain.ScaleHeight - 2180
        frmMain.WB.Left = 4180
        frmMain.WB.Height = frmMain.ScaleHeight - 2145
        frmMain.WB.Width = 11100
    
End If

End Sub
