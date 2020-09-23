VERSION 5.00
Begin VB.Form frmsearch 
   BackColor       =   &H00A97052&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search  The Web..."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00A97052&
      Caption         =   "Search Engines"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   410
      Width           =   5655
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Webcrawler"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   7
         Left            =   2730
         TabIndex        =   19
         Top             =   360
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "InfoSeek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   8
         Left            =   2730
         TabIndex        =   18
         Top             =   720
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "DejaNews"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   9
         Left            =   2730
         TabIndex        =   17
         Top             =   1080
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Inktomi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   10
         Left            =   2730
         TabIndex        =   16
         Top             =   1440
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Alta Vista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   3
         Left            =   1260
         TabIndex        =   15
         Top             =   360
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Open Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   4
         Left            =   1260
         TabIndex        =   14
         Top             =   720
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Lycos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   5
         Left            =   1260
         TabIndex        =   13
         Top             =   1080
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Excite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   6
         Left            =   1260
         TabIndex        =   12
         Top             =   1440
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Point"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1110
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Magellan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Yahoo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Point"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "Hotbot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "LookSmart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A97052&
         Caption         =   "AOL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8E9EC&
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yahoo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmsearch.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   750
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inktomi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         MouseIcon       =   "frmsearch.frx":06DC
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DejaNews"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         MouseIcon       =   "frmsearch.frx":082E
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InfoSeek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         MouseIcon       =   "frmsearch.frx":0980
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AlltheWeb"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         MouseIcon       =   "frmsearch.frx":0AD2
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         MouseIcon       =   "frmsearch.frx":0C24
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tapuz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         MouseIcon       =   "frmsearch.frx":0D76
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         MouseIcon       =   "frmsearch.frx":0EC8
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Webcrawler"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3060
         MouseIcon       =   "frmsearch.frx":101A
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lycos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmsearch.frx":116C
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1110
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Google"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmsearch.frx":12BE
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MSN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmsearch.frx":1410
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hotbot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4560
         MouseIcon       =   "frmsearch.frx":1562
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LookSmart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4560
         MouseIcon       =   "frmsearch.frx":16B4
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AOL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4560
         MouseIcon       =   "frmsearch.frx":1806
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1080
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A97052&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2340
      Width           =   5655
      Begin VB.TextBox txtSerchKey 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   360
         Width           =   2955
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Search"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search &For:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1680
      Picture         =   "frmsearch.frx":1958
      Stretch         =   -1  'True
      Top             =   30
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00A97052&
      Caption         =   "Search the Web"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   35
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-//=============================================================================
'-// Declare the required local variables
'-//=============================================================================
Dim Selected As Integer, Buffer As String

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

Private Sub Label1_Click()
Option1(1).Value = True
End Sub

Private Sub Label10_Click()
Option1(6).Value = True
End Sub

Private Sub Label11_Click()
Option1(5).Value = True
End Sub

Private Sub Label12_Click()
Option1(4).Value = True
End Sub

Private Sub Label13_Click()
Option1(3).Value = True
End Sub

Private Sub Label14_Click()
Option1(10).Value = True
End Sub

Private Sub Label15_Click()
Option1(13).Value = True
End Sub

Private Sub Label16_Click()
Option1(14).Value = True
End Sub

Private Sub Label3_Click()
Option1(8).Value = True
End Sub

Private Sub Label4_Click()
Option1(9).Value = True
End Sub

Private Sub Label5_Click()
Option1(2).Value = True
End Sub

Private Sub Label6_Click()
Option1(0).Value = True
End Sub

Private Sub Label7_Click()
Option1(11).Value = True
End Sub

Private Sub Label8_Click()
Option1(12).Value = True
End Sub

Private Sub Label9_Click()
Option1(7).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
      Selected = Index
      txtSerchKey.SetFocus
End Sub

Private Sub txtSerchKey_Change()

'-//=============================================================================
'-// Get the txtSerchKey TextBox text
'-//=============================================================================
    
      Buffer = txtSerchKey.Text
    
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

'-//=============================================================================
'-// Check if it is empty
'-//=============================================================================
        
        If Buffer = "" Then
        
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

'-//=============================================================================
'-// If so, disable the Search CommandButton
'-//=============================================================================
        cmdSearch.Enabled = False
    
   Else
'-//=============================================================================
'-// If not, enable the Search CommandButton
'-//=============================================================================
        cmdSearch.Enabled = True
        
    End If
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

End Sub

Private Sub cmdSearch_Click()
On Error GoTo ErrorHand
    Dim S As String

'-//===================================================================================
'-// GENERATE THE SEARCH COMMAND STRING FOR THE SELECTED
'-// SEARCH ENGINE
'-//====================================================================================

    Select Case Selected
    Case 0  'GOOGLE
         S = "http://www.google.com/custom?q=" & Buffer
         
    Case 1  'YAHOO!
        S = "http://search.yahoo.com/bin/search?p=" & Buffer
        
    Case 2 'LYCOS
        S = "http://sjc-search.sjc.lycos.com/default.asp?lpv=1&loc=searchhp&tab=web&query=" & Buffer
        
    Case 3  'AlltheWeb
        S = "http://alltheweb.com/search?cat=web&cs=utf8&q=" & Buffer & "&rys=0&_sb_lang=pref"
        
    Case 4 'The Web
          S = "http://www.lycos.com/cgi-bin/pursuit?query=" & Buffer + "&backlink=217&maxhits=25"
          
    Case 5 'Tapuz
        S = "http://www.tapuz.co.il/index/proceed.asp?q=" & Buffer
      
    Case 6 'EXCITE
        S = "http://www.excite.com/search.gw?searchType=Concept&search=" & Buffer & "&category=default&mode=relevance&showqbe=1&display=html3,hb"
        
    Case 7 'WEBCRAWLER
        S = "http://www.webcrawler.com/cgi-bin/WebQuery?searchText=" & Buffer & "&maxHits=25"
        
    Case 8 'INFOSeek
        S = "http://infoseek.go.com/Titles?col=WW&qt=%22" & Buffer & "%22&sv=IS&lk=noframes&svx=sbox_top&cc=WW&oq=" & Buffer
        
    Case 9  'DEJA
        S = "http://search.dejanews.com/nph-dnquery.xp?query=" & Buffer & "&defaultOp=AND&svcclass=dncurrent&maxhits=25"
        
    Case 10 'Inktomi
        S = "http://204.161.74.8:1234/query/?query=" & Buffer & "&hits=25&disp=Text+Only"
        
    Case 11 'MSN
        S = "http://search.msn.com/results.aspx?FORM=MSNH&q=" & Buffer
        
    Case 12 'Hotbot
        S = "http://www.hotbot.com/default.asp?prov=Inktomi&query=" & Buffer & "&ps=&loc=searchbox&tab=web"
        
    Case 13 'LookSmart
        S = "http://www.looksmart.com/r_search?l&inethome&pin=021228x584498b8690e88bb691&key=" & Buffer & "&search=0"
        
    Case 14 'AOL
        S = "http://search.aol.com/dirsearch.adp?start=&from=topsearchbox.%2Findex.adp&query=" & Buffer
    End Select
 
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

    Unload Me '-// UNLOAD THE FORM
    
'-//=========================================================================================
'-//Open the default Web Browser window with the selected location
'-//=========================================================================================

 frmMain.cboAddress = S
 frmMain.WB.Navigate (S)

'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

    Exit Sub
ErrorHand:
    MsgBox err.Description, vbOKOnly + vbInformation, "Unknown Error"
End Sub

