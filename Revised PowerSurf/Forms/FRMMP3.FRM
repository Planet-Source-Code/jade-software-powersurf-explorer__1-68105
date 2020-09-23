VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmmp3 
   BackColor       =   &H00E3F9FB&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3F9FB&
      Caption         =   "FILE  LIST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4605
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3870
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4170
         Left            =   1680
         Pattern         =   "*.midi;*.mid;*.mp3;*.wav;*.snd;*.wma;*.asx;*.m3u;*.asf;*.wma;*.cda"
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   2640
      Top             =   3000
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3F9FB&
      Caption         =   "NOW PLAYING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E3F9FB&
      Caption         =   "Mute ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMP3.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMP3.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMP3.frx":0934
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMP3.frx":0CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMP3.frx":1068
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpCapBor 
      BackColor       =   &H00F2DCCC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image imgpause 
      Height          =   240
      Left            =   600
      MouseIcon       =   "FRMMP3.frx":1602
      MousePointer    =   99  'Custom
      Picture         =   "FRMMP3.frx":1754
      ToolTipText     =   "Pause"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgstop 
      Height          =   240
      Left            =   360
      MouseIcon       =   "FRMMP3.frx":1CDE
      MousePointer    =   99  'Custom
      Picture         =   "FRMMP3.frx":1E30
      ToolTipText     =   "Stop"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgplay 
      Height          =   240
      Left            =   120
      MouseIcon       =   "FRMMP3.frx":23BA
      MousePointer    =   99  'Custom
      Picture         =   "FRMMP3.frx":250C
      ToolTipText     =   "Play"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PowerSurf WinMedia FX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   50
      Width           =   3585
   End
   Begin VB.Shape shpMB 
      BorderColor     =   &H00926747&
      BorderStyle     =   3  'Dot
      Height          =   6225
      Left            =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   3680
      MouseIcon       =   "FRMMP3.frx":2A96
      MousePointer    =   99  'Custom
      Picture         =   "FRMMP3.frx":2BE8
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   240
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMA 
      Height          =   2475
      Left            =   3960
      TabIndex        =   7
      Top             =   6240
      Width           =   3735
      URL             =   ""
      rate            =   1
      balance         =   5
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6588
      _cy             =   4366
   End
End
Attribute VB_Name = "frmmp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
