VERSION 5.00
Begin VB.Form frmoptions 
   BackColor       =   &H00A97052&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frmoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      ItemData        =   "frmoptions.frx":058A
      Left            =   120
      List            =   "frmoptions.frx":0597
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame frmblocked 
      BackColor       =   &H00A97052&
      Caption         =   "Definitions - Blocked Sites"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   2400
      TabIndex        =   1
      Top             =   30
      Width           =   3735
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ListBox LstURLDisplay 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1665
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   810
         Left            =   240
         Picture         =   "frmoptions.frx":05BD
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   810
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Remove Site"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "Remove a website that is banned. O.K it for viewing."
         Top             =   3800
         Width           =   1095
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add Site"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Add a new website to be banned."
         Top             =   3800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. :   google"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmoptions.frx":0CC4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   795
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame frmePrivacy 
      BackColor       =   &H00A97052&
      Caption         =   "Privacy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   2400
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox ChkFav 
         BackColor       =   &H00A97052&
         Caption         =   "Clear Favorites"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton cmddoit 
         Caption         =   "DO It!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   2280
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   120
         TabIndex        =   23
         Top             =   2010
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "&"
         TabIndex        =   22
         Top             =   3090
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "&"
         TabIndex        =   21
         Top             =   3450
         Width           =   2175
      End
      Begin VB.CommandButton cmdchangepass 
         Caption         =   "&Change Lock Password"
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   3810
         Width           =   2175
      End
      Begin VB.CheckBox chkurl 
         BackColor       =   &H00A97052&
         Caption         =   "Clear URL Visited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chkCookies 
         BackColor       =   &H00A97052&
         Caption         =   "Clear Cookies"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- BLOCKED/LOCK APPLICATION -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3090
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3450
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "You want to change FILTERING and LOCK APPLICATION Password ? "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   3495
      End
   End
   Begin VB.Frame frmstartup 
      BackColor       =   &H00A97052&
      Caption         =   "Start-Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   2400
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtdefine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
      Begin VB.OptionButton optDefine 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1695
      End
      Begin VB.OptionButton Opthome 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at Homepage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton Optgoogle 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at Google"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   3015
      End
      Begin VB.OptionButton OptBlank 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at Blank Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton OptLast 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at Last open site"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton OptYahoo 
         BackColor       =   &H00A97052&
         Caption         =   "Start PowerSurf at Yahoo!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "This will run on Start-up when program run."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAdd_Click()

If Text1 = "" Then MsgBox "Invalid Entry. Please check it!", vbCritical, "ERROR": Text1.SetFocus: Exit Sub

'-//============================================================================================
'-// CHECK IF EXTENSION NAME EXIST
'-//============================================================================================
If InStr(1, Text1, ".") > 0 Then MsgBox "Invalid Input. Please remove the extension name", vbCritical: Text1.SetFocus: Exit Sub

frmVerify.Show vbModal

End Sub

Private Sub cmdchangepass_Click()

If Text2 = "" And Text3 = "" Then MsgBox "Empty Fields. Please check it!", vbCritical: Text2.SetFocus: Exit Sub

'-//==================================================================================
'-// CHANGE LOCK PASSWORD
'-//==================================================================================

    Set rschangePass = New ADODB.Recordset
    rschangePass.Open "SELECT * FROM TBL_LOCK WHERE PASS='" & Text2.Text & "'", CN, adOpenStatic, adLockOptimistic
    
        If Not rschangePass.EOF Then
        
            rschangePass!Pass = Text3
            rschangePass.Update
            rschangePass.Requery
            
            MsgBox "Password successfully changed!", vbInformation
            
        
        ElseIf rschangePass.EOF Then
        
            MsgBox "No Password to change. Set-up new password by" & vbCrLf _
            & "clicking Lock Application on Tools Menu.", vbCritical
            Unload Me
            
        Else
            
            MsgBox "Old Password did not match.", vbCritical
            Text2 = ""
            Text3 = ""
            Text2.SetFocus
            Exit Sub
            
        End If
        
        '-//============================================================================
        '-// TO CLEAR VARIABLE FROM COMPUTER MEMORY
        '-//============================================================================
         Set rschangePass = Nothing

End Sub

Private Sub cmddoit_Click()

On Error Resume Next
Screen.MousePointer = vbHourglass

If chkCookies.Value = 1 Then Shell App.Path & "/BatchFiles/delcookies.bat", vbHide
If ChkFav.Value = 1 Then Kill App.Path & "/Favorites.ini"

If chkurl.Value = 1 Then

    '-//===============================================================================================
    '-// DELETE ALL URL VISITED FROM DATABASE
    '//================================================================================================
        CN.Execute "DELETE * FROM TBL_AUTOCOMPLETE"
    
    
    Set rsReloadURL = New ADODB.Recordset
    rsReloadURL.Open "SELECT * FROM TBL_AUTOCOMPLETE", CN, adOpenStatic, adLockOptimistic
    
    Do Until rsReloadURL.EOF
        frmMain.cboAddress.AddItem rsReloadURL!URL
        rsReloadURL.MoveNext
    Loop
    
        If rsReloadURL.RecordCount = 0 Then
            frmMain.cboAddress.Clear
            frmMain.cboAddress.Refresh
            frmMain.cboAddress = frmMain.WB.LocationURL
        End If
    
    Set rsReloadURL = Nothing
    
 End If
 
    Screen.MousePointer = vbDefault
    frmMain.WB.Refresh

End Sub

Private Sub CmdRemove_Click()

If Text1 = "" Then
MsgBox "Please Select item to remove.", vbCritical, "ERROR"
    LstURLDisplay.SetFocus
    Exit Sub
    
Else

    frmverifyremove.Show vbModal

End If


End Sub

Private Sub Form_Load()

 Dim Def, W, Sil As Integer
    
    Def = GetSetting("PowerSurf", "SKIN", "Default", Def)
    W = GetSetting("PowerSurf", "SKIN", "Wooden", W)
    Sil = GetSetting("PowerSurf", "SKIN", "Silver", Sil)
        
    If Def = 1 Then
        ModSkin.DefaultSkin
    ElseIf W = 1 Then
        ModSkin.WoodenSkin
    ElseIf Sil = 1 Then
    
        ModSkin.SilverGraySkin
    Else
    '-// CALL DEFAULT SKIN
        DefaultSkin
    End If
    
'-//=====================================================================================
'-// LOAD ALL BLOCKED WEBSITE TO LISTBOX
'-//=====================================================================================

Set RSLoadBlock = New ADODB.Recordset
RSLoadBlock.Open "SELECT * FROM TBL_BLOCKED_WEBSITES", CN, adOpenStatic, adLockOptimistic
    
With RSLoadBlock

    If .RecordCount > 0 Then .MoveFirst
        
    Do Until RSLoadBlock.EOF
        
            LstURLDisplay.AddItem RSLoadBlock!urladd
            RSLoadBlock.MoveNext
    
    Loop
    
   If .RecordCount > 0 Then .MoveLast
   
   .Requery
   LstURLDisplay.Refresh
    
 End With
 
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

'-//=====================================================================================
'-// GET VALUE FROM THE REGISTRY
'-//=====================================================================================

 Opthome = GetSetting("PowerSurf", "Settings", "Home", Opthome)
 OptLast = GetSetting("PowerSurf", "Settings", "Last", OptLast)
 OptBlank = GetSetting("PowerSurf", "Settings", "Blank", OptBlank)
 Optgoogle = GetSetting("PowerSurf", "Settings", "Google", Optgoogle)
 OptYahoo = GetSetting("PowerSurf", "Settings", "Yahoo", OptYahoo)
 txtdefine = GetSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
 optDefine = GetSetting("PowerSurf", "Settings", "DEFINE", optDefine)
End Sub

Private Sub Form_Unload(Cancel As Integer)

'-//======================================================================================================================
'-// CLEAR VARIABLES FROM COMPUTER'S MEMORY
'-//======================================================================================================================
      Set rsBlocksite = Nothing
      Set RSLoadBlock = Nothing
      Set RSDelBlock = Nothing
  
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

End Sub

Private Sub List1_Click()

If List1 = "Start-Up" Then

    frmstartup.Visible = True
    frmblocked.Visible = False
    frmePrivacy.Visible = False
    
ElseIf List1 = "Blocked Sites" Then
     
     frmblocked.Visible = True
     frmstartup.Visible = False
     frmePrivacy.Visible = False
   
ElseIf List1 = "Privacy" Then
     frmePrivacy.Visible = True
     frmstartup.Visible = False
     frmblocked.Visible = False
     
End If

End Sub

Private Sub LstURLDisplay_Click()
Text1 = LstURLDisplay.List(LstURLDisplay.ListIndex)
End Sub

Private Sub OptBlank_Click()

If OptBlank.Value = True Then
        Call SaveSetting("PowerSurf", "Settings", "Blank", 1)
        Call SaveSetting("PowerSurf", "Settings", "Home", 0)
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
        Call SaveSetting("PowerSurf", "Settings", "Google", 0)
        Call SaveSetting("PowerSurf", "Settings", "Yahoo", 0)
        Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
         txtdefine = ""
        Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
    Else
        Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
    End If
End Sub

Private Sub optDefine_Click()
    
If txtdefine = "" Then
        MsgBox "Please type the URL first.", vbCritical
        optDefine.Value = False
        Exit Sub
        
  Else
    
        If optDefine.Value = True Then
                 Call SaveSetting("PowerSurf", "Settings", "DEFINE", 1)
                 Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
               
                Call SaveSetting("PowerSurf", "Settings", "Google", 0)
                Call SaveSetting("PowerSurf", "Settings", "Home", 0)
                Call SaveSetting("PowerSurf", "Settings", "Last", 0)
                Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
                Call SaveSetting("PowerSurf", "Settings", "Yahoo", 0)
        Else
                Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
        End If
        
 End If
End Sub

Private Sub Optgoogle_Click()
If Optgoogle.Value = True Then
        Call SaveSetting("PowerSurf", "Settings", "Google", 1)
         Call SaveSetting("PowerSurf", "Settings", "Home", 0)
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
        Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
        Call SaveSetting("PowerSurf", "Settings", "Yahoo", 0)
        Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
        txtdefine = ""
        Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
    Else
        
        Call SaveSetting("PowerSurf", "Settings", "Google", 0)
    End If
End Sub

Private Sub optHome_Click()
    
    If Opthome.Value = True Then
        Call SaveSetting("PowerSurf", "Settings", "Home", 1)
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
        Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
        Call SaveSetting("PowerSurf", "Settings", "Google", 0)
        Call SaveSetting("PowerSurf", "Settings", "Yahoo", 0)
        Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
        txtdefine = ""
        Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
        
    Else
        Call SaveSetting("PowerSurf", "Settings", "Home", 0)
    End If


End Sub

Private Sub OptLast_Click()

    If OptLast.Value = True Then
        Call SaveSetting("PowerSurf", "Settings", "Last", 1)
        Call SaveSetting("PowerSurf", "SettingsVal", "Last", frmMain.WB.LocationURL)
        
        Call SaveSetting("PowerSurf", "Settings", "Home", 0)
        Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
        Call SaveSetting("PowerSurf", "Settings", "Google", 0)
        Call SaveSetting("PowerSurf", "Settings", "Yahoo", 0)
        Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
         txtdefine = ""
        Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
    Else
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
    End If
    
End Sub

Private Sub OptYahoo_Click()

    If OptYahoo.Value = True Then
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
        Call SaveSetting("PowerSurf", "SettingsVal", "Last", frmMain.WB.LocationURL)
        
        Call SaveSetting("PowerSurf", "Settings", "Home", 0)
        Call SaveSetting("PowerSurf", "Settings", "Blank", 0)
        Call SaveSetting("PowerSurf", "Settings", "Google", 0)
        Call SaveSetting("PowerSurf", "Settings", "Yahoo", 1)
        Call SaveSetting("PowerSurf", "Settings", "DEFINE", 0)
         txtdefine = ""
        Call SaveSetting("PowerSurf", "SettingsVal", "DEFINE", txtdefine)
    Else
        Call SaveSetting("PowerSurf", "Settings", "Last", 0)
    End If
    
End Sub
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/
