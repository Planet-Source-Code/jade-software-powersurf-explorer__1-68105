VERSION 5.00
Begin VB.Form frmlock 
   BackColor       =   &H00A97052&
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00A97052&
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   -80
      Width           =   5655
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   200
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00A97052&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmlock"
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLock = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

    Set rsLock = New ADODB.Recordset
    rsLock.Open "SELECT * FROM TBL_LOCK WHERE PASS='" & Text1 & "'", CN, adOpenStatic, adLockOptimistic
    
    If Not rsLock.EOF Then
        frmMain.Enabled = True
        Unload Me
    Else
    
        MsgBox "Invalid Password. Please try again.", vbCritical
        Text1 = ""
        Text1.SetFocus
        Exit Sub
        
    End If
        
End If

End Sub
