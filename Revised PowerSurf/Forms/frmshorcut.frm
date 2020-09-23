VERSION 5.00
Begin VB.Form frmshorcut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shorcut Keys"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmshorcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmshorcut.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmshorcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
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
    
End Sub
