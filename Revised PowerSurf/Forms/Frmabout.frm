VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About PowerSurf Explorer"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "Frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         Height          =   810
         Left            =   240
         Picture         =   "Frmabout.frx":0442
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "By: RJ GALLERMO © 2007"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PowerSurf Explorer (FREEWARE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   2385
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1200
         Picture         =   "Frmabout.frx":0C7E
         Top             =   840
         Width           =   480
      End
      Begin VB.Label LabEmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email PowerSurf  People."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   1800
         MouseIcon       =   "Frmabout.frx":1548
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Open your email client and lets you send email to PowerSurf Explorer."
         Top             =   1080
         Width           =   2160
      End
      Begin VB.Label LabHomePage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Visit PowerSurf Explorer Homepage."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   1800
         MouseIcon       =   "Frmabout.frx":210A
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Open your default browser and displays the PowerSurf Explorer homepage."
         Top             =   840
         Width           =   3120
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5895
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frmabout.frx":2CCC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2295
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   5760
         Y1              =   3580
         Y2              =   3580
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   5740
         X2              =   5740
         Y1              =   270
         Y2              =   3600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   30
         X2              =   30
         Y1              =   270
         Y2              =   3540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   2
         X1              =   60
         X2              =   5700
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00A56E3A&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   3345
         Left            =   0
         Top             =   240
         Width           =   5745
      End
   End
End
Attribute VB_Name = "frmabout"
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

Private Sub LabHomePage_Click()

    Shell ("explorer http://URL.com"), vbNormalNoFocus
    LabHomePage.ForeColor = vbYellow
    
End Sub

Private Sub LabEmail_Click()

   Dim PersonName, Subject As String
  
   PersonName = "admin@powersurf.tk"
   Subject = "Comments and Bugs"
   ShellExecute hWnd, "open", "mailto:" & PersonName & "?subject=" & Subject, "", "", vbNormalFocus
   LabEmail.ForeColor = vbYellow
   
End Sub

Private Sub LabEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
LabEmail.ForeColor = vbRed
LabHomePage.ForeColor = vbYellow
End Sub

Private Sub LabHomePage_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
LabHomePage.ForeColor = vbRed
LabEmail.ForeColor = vbYellow
End Sub

