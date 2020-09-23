VERSION 5.00
Begin VB.Form frmViewFav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Favorites"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmViewFav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   4125
      Left            =   120
      MouseIcon       =   "frmViewFav.frx":038A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmViewFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    '-//===========================================================================================
    '-// LOAD FAVORITES
    '-//===========================================================================================

    Dim sText As String
    Dim X As Integer
    X = FreeFile
    
    On Error GoTo R
    
    Open App.Path & "/Favorites.ini" For Input As #X
    While Not EOF(X)
        Input #X, sText$
            List1.AddItem sText$
            DoEvents
    Wend
    Close #X
    Exit Sub
    
R:
    If err.Number = 53 Then Exit Sub
    
End Sub

Private Sub List1_DblClick()
frmMain.cboAddress = List1
frmMain.WB.Navigate frmMain.cboAddress
Unload Me
End Sub
