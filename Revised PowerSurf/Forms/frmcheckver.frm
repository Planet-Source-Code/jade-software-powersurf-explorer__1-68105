VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmcheckver 
   BorderStyle     =   0  'None
   Caption         =   "Agent Check for Updates"
   ClientHeight    =   90
   ClientLeft      =   18015
   ClientTop       =   13590
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3960
      Top             =   1680
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3960
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3960
      Top             =   0
   End
End
Attribute VB_Name = "frmcheckver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As IAgentCtlCharacterEx
Dim Status As String
Dim File As String

Private Sub Form_Load()
    Agent1.Characters.Load "ID"
    Set c = Agent1.Characters.Character("ID")
    c.Show
    c.Speak "Please be patient while checking for new updates."
    
    '=============================================================================
'IMPORTANT: enter you website path in place of these local paths
'=============================================================================
    Status = Inet1.OpenURL("file://" & App.Path & "\status.txt")
    File = Inet2.OpenURL("file://" & App.Path & "\file.txt")
    '=========================================================================
    'for example your site is http://microsoft.com
    'then type "http://microsoft.com/statusfile.txt"
    'and "http://microsoft.com/file.txt"
    '=========================================================================
    Timer1.Enabled = True
    
End Sub

Private Sub Agent1_Bookmark(ByVal BookmarkID As Long)
    On Error Resume Next
    Select Case BookmarkID
    Case 1
        If MsgBox("Would you like to download it now!", vbQuestion + vbYesNo, "Download Now!") = vbYes Then
            Launch
        End If
    Case 2
        If MsgBox("sorry no updates found, Would you like to visit the product home page now?", vbQuestion + vbYesNo, "Download Now!") = vbYes Then
            Launch2
        End If
    End Select
End Sub

Private Sub Timer1_Timer()
    If Status = "1" Then
        c.StopAll
        c.Speak "Congratulations, new version of PowerSurf Explorer is available."
    Else
        c.StopAll
        c.Speak "Sorry, no update available"
    End If
End Sub

Sub Launch()
    ShellExecute Me.hWnd, "open", File, "", 0, 1
End Sub

Sub Launch2()
    ShellExecute Me.hWnd, "open", "http://msagentworld.tripod.com", "", 0, 1
End Sub
