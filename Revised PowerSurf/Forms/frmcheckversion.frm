VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmcheckversion 
   BackColor       =   &H00A56E3A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check New Version!"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2280
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   230
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdcheck 
      Caption         =   "Check Now!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   235
      Left            =   2040
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "What's  New!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblnew 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0"
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
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblcurr 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0"
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
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Version:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Version:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmcheckversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcheck_Click()

PB.Value = 0

PB.Visible = True
Timer1.Enabled = True
cmdcheck.Enabled = False

lblnew = ""
Text2 = ""

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
    
   Me.Icon = frmMain.ImgList.ListImages(16).ExtractIcon
lblcurr = App.Major & "." & App.Minor

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim ans
                  
    ans = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Check New Version!")
                 
    If ans = vbYes Then
        Unload Me
    Else
        Cancel = 1 '-// TO PREVENT FROM CLOSING THE FORM BY PRESSING (X) BUTTON
    End If
    
End Sub

Private Sub Timer1_Timer()

PB.Value = PB.Value + 1

If PB.Value = 8 Then lblnew = "Checking for updates..."

    If PB.Value = 100 Then
     
        '-//=================================================================================
        '-// GRABBED PROGRAM VERSION
        '-//=================================================================================
        lblnew = Inet1.OpenURL("http://www.angelfire.com/planet/powersurf/progver.ver")
            
        '-//=================================================================================
        '-// COMPARE CURRENT VERSION TO THE INTERNET
        '-//=================================================================================
        If lblnew > App.Major & App.Minor Then
        
        '-//=================================================================================
        '-// GRABBED WHAT'S NEW
        '-//=================================================================================
            Text2 = Inet1.OpenURL("http://www.angelfire.com/planet/powersurf/whatsnew.ver")
            MsgBox "There is an update available!", vbInformation
                  
                  Dim response
                  
                    response = MsgBox("Download File?", vbQuestion + vbYesNo, "Download Now!")
                 
                   If response = vbYes Then
                   '-//=================================================================================
                   '-// DOWNLOAD FILE IF NEW VERSION FOUND
                   '-//=================================================================================
                        DoFileDownload StrConv("http://www.angelfire.com/planet/powersurf/powerfile.zip", vbUnicode)
                   End If
                    
        Else
            
            MsgBox "You have the most up-to-date version available", vbInformation
            lblnew = "No Updates available."
            Text2 = lblnew
            
        End If
               
        Timer1.Enabled = False
        PB.Visible = False
        cmdcheck.Enabled = True
        
    End If
    

    
End Sub
