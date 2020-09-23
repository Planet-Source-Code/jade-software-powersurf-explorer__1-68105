VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmsource 
   BackColor       =   &H00A97052&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Source"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8295
   Icon            =   "frmsource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTBox 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12726
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmsource.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuopen 
         Caption         =   "Open...         "
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmsource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-//===========================================================================
'-// VIEW SOURCE
'-//============================================================================

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
    
    On Error Resume Next
     RTBox = frmMain.WB.Document.documentElement.innerHTML

End Sub
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuopen_Click()
Dim sFile As String
'-//===========================================================================
'-// FILTERING FILE TYPES TO OPEN
'-//============================================================================

    With CDL
        .DialogTitle = "Open"
        .Filter = "TEXT FILES (*.txt)|*.txt|HTML FILES (*.htm;*.html)|*.htm;*.html"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    RTBox.LoadFile sFile
    Me.Caption = sFile
End Sub
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

Private Sub mnuSaveas_Click()
Dim sFile As String

    With CDL
        .DialogTitle = "Save As"
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        sFile = .FileName
    End With
   Me.Caption = sFile
    RTBox.SaveFile sFile
    
End Sub
