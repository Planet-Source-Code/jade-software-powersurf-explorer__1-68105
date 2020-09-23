VERSION 5.00
Begin VB.Form frmsetpassword 
   BackColor       =   &H00A97052&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Password"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdset 
      Caption         =   "&Set Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "@"
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A97052&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4095
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
         Left            =   1680
         MaxLength       =   8
         PasswordChar    =   "@"
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmsetpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdset_Click()

If Trim(Text1) = "" And Trim(Text2) = "" Then MsgBox "Empty fields. Please check it!", vbCritical: Text1.SetFocus: Exit Sub

If Text1 = Text2 Then

    Set rsAddPass = New ADODB.Recordset
    rsAddPass.Open "SELECT * FROM TBL_LOCK", CN, adOpenStatic, adLockOptimistic
            
        '-//============================================================================
        '-// ADDING NEW PASSWORD TO DATABASE
        '-//============================================================================
             If rsAddPass.EOF Then
        
            With rsAddPass
                .AddNew
                !Pass = Text2 & "" '-// INVALID USE OF NULL
                .Update
                .Requery
            End With
        
        Set rsAddPass = Nothing
        
        '-//============================================================================
        '-// DISABLE ALL CONTROL AFTER SAVING RECORDS, TO AVOID ADDING AGAIN AND
        '-// LOAD LOCK APPLICATION SYSTEM
        '-//============================================================================
        
            cmdset.Enabled = False
            cmdcancel.Enabled = False
            Text1.Locked = True
            Text2.Locked = True
            
            MsgBox "New Password succesfully set!", vbInformation
            Unload Me
            frmlock.Show vbModal
        
        End If
Else
    MsgBox "Password you entered did not match.", vbCritical
    Exit Sub
End If

End Sub

Private Sub Form_Load()

'-//=====================================================================================
'-// GRABBED ICON FROM IMAGELIST
'-//=====================================================================================
Me.Icon = frmMain.ImgList.ListImages(14).ExtractIcon

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
    
Text1 = ""
Text2 = ""

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Text1 <> "" Then
        Text2.SetFocus
    Else
        MsgBox "Invalid input. Please check it!", vbCritical
        Exit Sub
    End If
    
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Text2 <> "" Then
        cmdset.SetFocus
    Else
        MsgBox "Invalid input. Please check it!", vbCritical
        Exit Sub
    End If
End If

End Sub
