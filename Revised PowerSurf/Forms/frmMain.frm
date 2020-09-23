VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H006B553C&
   BorderStyle     =   0  'None
   Caption         =   "PowerSurf Explorer v1.0"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10980
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7080
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "NewWindow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1C
            Key             =   "startup"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15F6
            Key             =   "OPenfile"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B90
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":212A
            Key             =   "PreviewPrint"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26C4
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C5E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31F8
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3792
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D2C
            Key             =   "All"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42C6
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4860
            Key             =   "find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DFA
            Key             =   "about"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5394
            Key             =   "WindowsUpdate"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":592E
            Key             =   "LockApp"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EC8
            Key             =   "Fav"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6262
            Key             =   "CheckVer"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65FC
            Key             =   "Personal"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B96
            Key             =   "options"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7130
            Key             =   "player"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76CA
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF2C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E4C6
            Key             =   "home"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA60
            Key             =   "search"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EFFA
            Key             =   "back"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F594
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB2E
            Key             =   "msn"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FEC8
            Key             =   "ViewSource"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B7E
            Key             =   "skin"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17126
            Key             =   "SystemReq"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":176CE
            Key             =   "ShortKey"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B22
            Key             =   "ym"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17E3E
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":183D8
            Key             =   "net"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18972
            Key             =   "explorer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18F0C
            Key             =   "startmenu"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":194A6
            Key             =   "programs"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picskin 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   6240
      ScaleHeight     =   3105
      ScaleWidth      =   2265
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ListBox List1 
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
         ForeColor       =   &H00FF0000&
         Height          =   2175
         ItemData        =   "frmMain.frx":19A40
         Left            =   120
         List            =   "frmMain.frx":19A4D
         MouseIcon       =   "frmMain.frx":19A73
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdskin 
         Caption         =   "Apply &Skin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Skin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   20
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2280
         X2              =   0
         Y1              =   220
         Y2              =   220
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   2030
         MouseIcon       =   "frmMain.frx":19BC5
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":19D17
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   255
      End
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   6735
      Left            =   80
      TabIndex        =   2
      Top             =   1800
      Width           =   15420
      ExtentX         =   27199
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imgTop16x16 
      Left            =   5280
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A157
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A6F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AC8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B225
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B7BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD59
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C2F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C88D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CE27
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D1C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D613
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D9AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DF47
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E4E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EA7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F015
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F16F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F709
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FCA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2023D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":207D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27039
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D89B
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DE35
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E575
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E3F9FB&
      Height          =   8715
      Left            =   360
      ScaleHeight     =   8655
      ScaleWidth      =   3975
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   4035
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
         Height          =   7125
         Left            =   120
         TabIndex        =   23
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
            Height          =   6390
            Left            =   120
            TabIndex        =   26
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
            Height          =   6690
            Left            =   1680
            Pattern         =   "*.midi;*.mid;*.mp3;*.wav;*.snd;*.wma;*.asx;*.m3u;*.asf;*.wma;*.cda"
            TabIndex        =   25
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
            TabIndex        =   24
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
         TabIndex        =   21
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
            TabIndex        =   22
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
         TabIndex        =   20
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
               Picture         =   "frmMain.frx":2EB0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2EEA9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F443
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F7DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2FB77
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin WMPLibCtl.WindowsMediaPlayer WMA 
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   5640
         Width           =   3135
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   0   'False
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
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
         _cx             =   5530
         _cy             =   1720
      End
      Begin VB.Image imgpause 
         Height          =   240
         Left            =   600
         MouseIcon       =   "frmMain.frx":30111
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":30263
         ToolTipText     =   "Pause"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgstop 
         Height          =   240
         Left            =   360
         MouseIcon       =   "frmMain.frx":307ED
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3093F
         ToolTipText     =   "Stop"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgplay 
         Height          =   240
         Left            =   120
         MouseIcon       =   "frmMain.frx":30EC9
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3101B
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
         TabIndex        =   27
         Top             =   50
         Width           =   3585
      End
      Begin VB.Shape shpMB 
         BorderColor     =   &H00926747&
         BorderStyle     =   3  'Dot
         Height          =   8745
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   3680
         MouseIcon       =   "frmMain.frx":315A5
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":316F7
         ToolTipText     =   "Close"
         Top             =   30
         Width           =   240
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
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   210
      Left            =   6720
      TabIndex        =   7
      Top             =   6480
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox pictop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10980
      TabIndex        =   4
      Top             =   0
      Width           =   10980
      Begin VB.Image Image1 
         Height          =   250
         Left            =   360
         Picture         =   "frmMain.frx":31C81
         Stretch         =   -1  'True
         Top             =   90
         Width           =   250
      End
      Begin VB.Image cmglose 
         Height          =   240
         Left            =   14955
         MouseIcon       =   "frmMain.frx":3254B
         MousePointer    =   99  'Custom
         ToolTipText     =   "Close"
         Top             =   105
         Width           =   240
      End
      Begin VB.Image imgmin 
         Height          =   240
         Left            =   14640
         MouseIcon       =   "frmMain.frx":3269D
         MousePointer    =   99  'Custom
         ToolTipText     =   "Minimize"
         Top             =   105
         Width           =   240
      End
      Begin VB.Label lblLocationname 
         BackStyle       =   0  'Transparent
         Caption         =   "LocationName"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   120
         Width           =   11295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PowerSurf Explorer "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   2295
      End
   End
   Begin ComCtl2.Animation Ani 
      Height          =   430
      Left            =   14975
      TabIndex        =   3
      Top             =   430
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   741
      _Version        =   327681
      AutoPlay        =   -1  'True
      BackColor       =   11104338
      FullWidth       =   49
      FullHeight      =   28
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   8595
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6244
            MinWidth        =   6244
         EndProperty
      EndProperty
   End
   Begin VBSmartXPMenu.SmartMenuXP MenuXP 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   465
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      FontForeColor   =   16777215
      FontBackColor   =   -2147483636
      CheckBackColor  =   192
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
   Begin MSComctlLib.Toolbar TBTOP 
      Height          =   540
      Left            =   120
      TabIndex        =   8
      Top             =   855
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   953
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "imgTop16x16"
      HotImageList    =   "imgTop16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Notepad"
            ImageIndex      =   20
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnuWindow"
                  Text            =   "Window"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnusep"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "contact"
                  Text            =   "Contact"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NetMeeting"
                  Text            =   "NetMeeting     "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Fast Searching..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            Key             =   "History"
            Object.ToolTipText     =   "History"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mail"
            ImageIndex      =   18
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnuCheckMail"
                  Text            =   "Check Mail      "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SendMail"
                  Text            =   "Send Mail    "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Text Size"
            Key             =   "textsize"
            Object.ToolTipText     =   "Text Size"
            ImageIndex      =   23
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnulargest"
                  Text            =   "&Largest    "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnularge"
                  Text            =   "Lar&ge"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnumedium"
                  Text            =   "&Medium"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnusmall"
                  Text            =   "&Small"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnusmallest"
                  Text            =   "Sma&llest"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Media"
            Key             =   "player"
            Object.ToolTipText     =   "WinMedia Player"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lock"
            Key             =   "lock"
            Object.ToolTipText     =   "Lock PowerSurf"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contents"
            Key             =   "Contents"
            Object.ToolTipText     =   "Contents"
            ImageIndex      =   5
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":327EF
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   960
      Left            =   0
      TabIndex        =   9
      Top             =   820
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   1693
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   12975
      _CBHeight       =   960
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   540
      Width1          =   9675
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Address"
      Child2          =   "cboAddress"
      MinHeight2      =   315
      Width2          =   7995
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Caption3        =   "VB Links"
      Child3          =   "cboURLList"
      MinHeight3      =   315
      Width3          =   3195
      FixedBackground3=   0   'False
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Caption4        =   "JavaScript Links"
      Child4          =   "cboJS"
      MinHeight4      =   315
      FixedBackground4=   0   'False
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Text            =   "Enter Website address"
         Top             =   600
         Width           =   7050
      End
      Begin VB.ComboBox cboURLList 
         Height          =   315
         ItemData        =   "frmMain.frx":32951
         Left            =   8985
         List            =   "frmMain.frx":32997
         TabIndex        =   11
         Text            =   "Select a VB Site"
         ToolTipText     =   "Select the dropdown box for quick shortcuts"
         Top             =   600
         Width           =   2205
      End
      Begin VB.ComboBox cboJS 
         Height          =   315
         ItemData        =   "frmMain.frx":32C83
         Left            =   12915
         List            =   "frmMain.frx":32CB4
         TabIndex        =   10
         Text            =   "Select a JavaScript Site"
         Top             =   600
         Width           =   390
      End
   End
   Begin SHDocVwCtl.WebBrowser WBFav 
      Height          =   6225
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
      ExtentX         =   4260
      ExtentY         =   10980
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComCtl3.CoolBar CoolBarFav 
      Height          =   390
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   688
      BandCount       =   1
      ForeColor       =   8388608
      BackColor       =   8438015
      EmbossHighlight =   4210752
      _CBWidth        =   2415
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Caption1        =   "History"
      MinHeight1      =   300
      Width1          =   2880
      NewRow1         =   0   'False
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2080
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Visible         =   0   'False
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PopUp As Boolean                '-// IF TRUE ALLOW, IF FALSE BLOCKED
Dim lleft, ltop, lwidth As Long     '-// FOR WIDTH, TOP, LEFT oF CONTROLS
Dim FavoriteAddress As String       '-// FOR ADDING FAVORITES

Private Sub cboAddress_Change()

Set rsAutoComplete = New ADODB.Recordset
rsAutoComplete.Open "SELECT * FROM TBL_autocomplete WHERE url='" & cboAddress.Text & "'", CN, adOpenStatic, adLockOptimistic

If Not rsAutoComplete.EOF Then
    cboAddress.SelStart = 0
    cboAddress.SelLength = Len(cboAddress.Text)
    WB.Navigate cboAddress
End If
Set rsAutoComplete = Nothing
End Sub

Private Sub cboAddress_Click()
    '-//========================================================================================
    '-// NAVIGATE
    '-//========================================================================================
    If cboAddress.Text <> "" Then WB.Navigate cboAddress.Text
End Sub

Private Sub cboJS_Click()

cboAddress = cboJS
    
    '-//========================================================================================
    '-// NAVIGATE
    '-//========================================================================================
    If cboAddress.Text <> "" Then WB.Navigate cboAddress.Text
    
    Set rsAutoComplete = New ADODB.Recordset
    rsAutoComplete.Open "SELECT * FROM TBL_AUTOCOMPLETE WHERE url='" & cboAddress.Text & "'", CN, adOpenStatic, adLockOptimistic
    
    If rsAutoComplete.RecordCount <> 1 Then
        With rsAutoComplete
        
            .AddNew
            !URL = cboAddress
            .Update
            .Requery
        
        End With
        cboAddress.AddItem cboAddress.Text
    End If
    
    Set rsAutoComplete = Nothing
    
End Sub

Private Sub cboURLList_Click()

cboAddress = cboURLList

   '-//=======================================================
    '-// NAVIGATE
    '-//=======================================================
    If cboAddress.Text <> "" Then WB.Navigate cboAddress.Text
    
    Set rsAutoComplete = New ADODB.Recordset
    rsAutoComplete.Open "SELECT * FROM TBL_AUTOCOMPLETE WHERE url='" & cboAddress.Text & "'", CN, adOpenStatic, adLockOptimistic
    
    If rsAutoComplete.RecordCount <> 1 Then
        With rsAutoComplete
        
            .AddNew
            !URL = cboAddress
            .Update
            .Requery
        
        End With
        cboAddress.AddItem cboAddress.Text
    End If
    
    Set rsAutoComplete = Nothing
 
End Sub

Sub Check1_Click()

If Check1.Value = 1 Then
    WMA.settings.mute = True
Else
    WMA.settings.mute = False
End If

End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path

End Sub



Private Sub Image5_Click()
        Picture1.Visible = False
        WB.Left = 75
        WB.Top = 1800
        WB.Height = Me.ScaleHeight - 2145
        WB.Width = Me.ScaleWidth - 150
End Sub

Private Sub imgpause_Click()
'-//==============================================
'-// PAUSED MUSIC
'-//==============================================
If WMA.playState = wmppsPlaying Then WMA.Controls.pause
Timer1.Enabled = False
End Sub

Private Sub imgplay_Click()

If File1.FileName = "" Then Exit Sub
WMA.URL = File1.Path & "\" & File1.FileName
WMA.Controls.play

lblname.Caption = "Now Playing: " & File1.FileName
Timer1.Enabled = True
End Sub

Private Sub imgstop_Click()
WMA.Controls.Stop
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
 
 If lblname.Left < -1000 Then
        lblname.Left = 7000
    Else
        lblname.Left = Val(lblname.Left) - 200
    End If

End Sub



Private Sub cmdClose_Click()

'-//=====================================================================================
'-// CLOSE HISTORY
'-//=====================================================================================
        WBFav.Visible = False
        CoolBarFav.Visible = False
        cmdClose.Visible = False
        
        WB.Left = 75
        WB.Top = 1800
        WB.Height = Me.ScaleHeight - 2145
        WB.Width = Me.ScaleWidth - 150
       
        
End Sub

Private Sub cmdskin_Click()

Screen.MousePointer = vbHourglass
Call ModSkin.SelectSkin(True)
Screen.MousePointer = vbDefault
Me.Refresh

MsgBox "Changes has been successful.", vbInformation
    
End Sub

Private Sub cmglose_Click()
End
End Sub

Private Sub Drive1_Change()
On Error GoTo errorhandler


File1.Path = Drive1.Drive
Dir1.Path = Drive1.Drive

Exit Sub

errorhandler:
 MsgBox "Device Unavailable", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub Form_Load()
  
    '-//=======================================================================================
    '-// GET CONNECT TO DATABASE
    '-//=======================================================================================
    
            get_connected
    
    pBuildMenus  '-// BUILD THE MAIN MENU
    
   
    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/
    
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
    
    lblLocationname = ""  '-//CLEAR LOCATION NAME
    
    '-//=======================================================================================
    '-// RESIZE AND POSITION FRMMAIN
    '-//=======================================================================================
    'Me.Height = 11100
    Me.Height = 11000
    Me.Top = 0
    Me.Width = 15360
  
    
'-//=======================================================================================
'-// LOAD AND PLAY LOGO (.AVI)
'-//=======================================================================================
  '-//=================================================================================
            '-// IT USES MICROSOFT SCRIPTING RUNTIME
            '-//================================================================================
            Dim fso1 As New FileSystemObject
                         
            If fso1.FileExists(App.Path & "/Videos/logo.avi") Then
            
            '-//==================================================================================
            '-// CALL VIDEO FILE, IF NOT EXIST THEN MSGBOX WILL BE DISPLAYED
            '-//==================================================================================
                Ani.Open App.Path & "/Videos/logo.avi"
                Ani.play
            Else
            
                MsgBox ("Video files not found. Program will now terminate!"), vbExclamation, "Not Found"
                End
            End If
    
    
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/
  
Set rsAdditem = New ADODB.Recordset
     rsAdditem.Open "SELECT * FROM TBL_AUTOCOMPLETE", CN, adOpenStatic, adLockOptimistic
    
    Do Until rsAdditem.EOF
        cboAddress.AddItem rsAdditem!URL
        rsAdditem.MoveNext
    Loop
    
Set rsAdditem = Nothing

  '-//==========================================================================================
  '-// POP-UP BLOCKER SETTINGS (REGISTRY)
  '-//==========================================================================================

  Dim S As String
    S = GetSetting("PowerSurf", "Option", "Popup", S)
    
    If S = "1" Then
                          
        PopUp = False
        MenuXP.MenuItems.Value("KeyPopblock") = smiChecked
        StatusBar.Panels(3) = "Enable Pop-Ups"
         Call SaveSetting("PowerSurf", "Option", "Popup", "1")
    Else
        PopUp = True
         MenuXP.MenuItems.Value("KeyPopblock") = smiUnchecked
        StatusBar.Panels(3) = "Disable Pop-Ups"
        Call SaveSetting("PowerSurf", "Option", "Popup", "0")
    End If
    
  '-//==========================================================================================
  '-// START-UP SETTINGS (REGISTRY)
  '-//==========================================================================================

On Error GoTo R
   
   Dim H, L, b, G, y, D As String
   
   H = GetSetting("PowerSurf", "Settings", "Home", H)
   L = GetSetting("PowerSurf", "Settings", "Last", L)
   b = GetSetting("PowerSurf", "Settings", "Blank", b)
   G = GetSetting("PowerSurf", "Settings", "Google", G)
   y = GetSetting("PowerSurf", "Settings", "Yahoo", y)
   D = GetSetting("PowerSurf", "Settings", "DEFINE", D)
   
   If D = 1 Then WB.Navigate GetSetting("PowerSurf", "SettingsVal", "DEFINE", WB.LocationURL)
   If G = 1 Then WB.Navigate "http://www.google.com"        '
   If H = 1 Then WB.GoHome
   If L = 1 Then WB.Navigate (GetSetting("PowerSurf", "SettingsVal", "Last", WB.LocationURL))
   If y = 1 Then WB.Navigate "www.Yahoo.com"
   If b = 1 Then WB.Navigate "about:blank"
  
  Exit Sub

    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/
R:
If err.Number = 13 Then WB.Navigate App.Path & "/Resources/Home/index.htm"


End Sub

Private Sub Form_Resize()

On Error Resume Next
'-//==============================================================================================
'-// RESIZE/ADJUST THE COOLBAR
'-//==============================================================================================
  
    Dim xRatio
    xRatio = (Me.ScaleWidth * 100) \ Me.ScaleWidth
        
        lleft = CLng((CoolBar1.Left * xRatio) \ 100)
        ltop = CoolBar1.Top
        lwidth = Me.ScaleWidth
        CoolBar1.Move lleft, ltop, lwidth
        
    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/

'-//=============================================================================================
'-// RESIZE/ADJUST WEB BROWSER CONTROL
'-//=============================================================================================

    WB.Height = Me.ScaleHeight - 2145
    WB.Width = Me.ScaleWidth - 150
    WB.Top = Me.ScaleTop + 1820

    
    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/
    
'-//=======================================================================================
'-// RESIZE/ADJUST THE STATUS BAR
'-//=======================================================================================
 
lwidth = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width)
        
        If lwidth > 0 Then
        
           StatusBar.Panels(1).Width = lwidth
           StatusBar.Panels(1).Visible = True
                 
        Else
        
           StatusBar.Panels(1).Visible = False
        
        End If
    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/

'-//=======================================================================================
'-// RESIZE/ADJUST THE PROGRESS BAR
'-//=======================================================================================

        PB.Left = Me.ScaleWidth - 5940
        PB.Top = Me.ScaleHeight - 225

    
End Sub

Private Sub Picture3_Click()
WB.Navigate cboAddress
End Sub

Private Sub Form_Unload(Cancel As Integer)

        Set rsBlock = Nothing
        Set rsAutoComplete = Nothing
         
End Sub

Private Sub Image2_Click()
picskin.Visible = False
End Sub

Private Sub imgmin_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub TBTOP_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key

    Case "Search"
        frmsearch.Show vbModal
    Case "Back"
        On Error Resume Next
        WB.GoBack
    Case "Forward"
        On Error Resume Next
        WB.GoForward
    Case "Stop"
        On Error Resume Next
        WB.Stop
    Case "Refresh"
        On Error Resume Next
        WB.Refresh
    Case "Home"
        On Error Resume Next
        WB.GoHome
        
    Case "Print"
        WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
    
    Case "lock"
            
            '-//==================================================================================
            '-// VERIFY IF PASSWORD FOR LOCK APPLICATION HAS BEEN SET
            '-// IF NOT FRMSETPASSWORD WILL SHOW UP
            '-//==================================================================================

            Set rschkPass = New ADODB.Recordset
            rschkPass.Open "SELECT * FROM TBL_LOCK", CN, adOpenStatic, adLockOptimistic
            
                If rschkPass.EOF Then
                    frmsetpassword.Show vbModal
                Else
                    frmlock.Show vbModal
                End If
                
            Set rschkPass = Nothing
    Case "player"
    
       
        hist = False
        cmdClose_Click
        
        sPlay = True
        viewplayer
        
    Case "History"
       
        hist = True
        ViewHistory
        
        sPlay = False
        Picture1.Visible = False
        
    Case "Contents"
            
            '-//=================================================================================
            '-// IT USES MICROSOFT SCRIPTING RUNTIME
            '-//================================================================================
            Dim fso As New FileSystemObject
                         
            If fso.FileExists(App.Path & "/powersurf.chm") Then
            
            '-//==================================================================================
            '-// CALL HELP FILE, IF NOT EXIST THEN MSGBOX WILL BE DISPLAYED
            '-//==================================================================================
            ShellExecute hWnd, "open", App.Path & "/powersurf.chm", "", "", vbNormalFocus
   
            Else
            
                MsgBox ("PowerSurf Help file not found."), vbCritical
                Exit Sub
            End If
            
End Select

End Sub

Private Sub TBTOP_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next

Select Case ButtonMenu.Key
    
    Case "NetMeeting"
        On Error GoTo ErrHandler

        Dim netmeeting As String
        netmeeting = Shell("C:\Program Files\NetMeeting\Conf.exe", vbNormalFocus)

        Exit Sub

ErrHandler:
        
        MsgBox "No installed Net Meeting on your computer." & vbCrLf & _
               "Or maybe installed in different location.", vbCritical
       
   
    Case "contact"
    
    On Error GoTo Errhand
    
        Dim contact
        contact = Shell("C:\Program Files\Outlook Express\WAB.EXE", vbNormalFocus)
        Exit Sub
Errhand:
        MsgBox "No installed Address Book on your computer." & vbCrLf & _
               "Or maybe installed in different location.", vbCritical
       
    Case "mnuWindow"
    
            Shell App.Path & "/PowerSurf.exe", vbNormalFocus
    
    Case "mnuCheckMail"
    
             On Error GoTo err

            Shell "C:\Program Files\Outlook Express\msimn.exe", vbNormalFocus
            
            Exit Sub
err:
            MsgBox "No installed Yahoo Messenger on your computer!" & vbCrLf & _
                   "Or maybe installed in different location.", vbCritical
    
    Case "SendMail"

            Dim Person, Subject As String
        
            Person = InputBox("Enter email address", "email")
            Subject = InputBox("Enter subject for email", "subject")
            
           ShellExecute hWnd, "open", "mailto:" & Person & "?subject=" & Subject, "", "", vbNormalFocus
                       
    Case "mnulargest"
        WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
    Case "mnularge"
        WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
    Case "mnumedium"
        WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
    Case "mnusmall"
        WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
    Case "mnusmallest"
        WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
        
End Select

End Sub

Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    On Error Resume Next
    '-//=============================================================================================
    '-// PROCEDURE TO BANNED WEBSITE
    '-//=============================================================================================
        Set rsBlock = New ADODB.Recordset
        rsBlock.Open "SELECT * FROM  TBL_BLOCKED_WEBSITES", CN, adOpenStatic, adLockOptimistic
    
               
        Dim i As Integer
    
    '-//=============================================================================================
    '-// MOVE TO FIRST RECORD
    '-//=============================================================================================

        rsBlock.MoveFirst
    
    '-//=============================================================================================
    '-// LOOP/SEARCH TO EACH RECORD IF EXIST THEN IF URL FOUND ON DATABASE
    '-// NAVIGATE TO BLOCK.HTM
    '-//=============================================================================================

        For i = 0 To rsBlock.RecordCount - 1
        
        '-// IF YOU WANT TO BANNED/BLOCKED A SITE
        '-// JUST ADD THE URL ADDRESS WITHOUT extension name like ".com"
            
            If InStr(1, URL, rsBlock!urladd) > 0 Then
                WB.Navigate App.Path & "/Resources/block.htm"
                Cancel = True '-// TO STOP/BLOCK OR EXIT SUB
            End If
            
            rsBlock.MoveNext
            
        Next
            
        
End Sub

Private Sub WB_NewWindow2(ppDisp As Object, Cancel As Boolean)

'-//=============================================================================================
'-// THIS WILL ALLOW A POP-UP TO BE LOAD OR TO BE BLOCKED!
'-//  CANCEL = MEANS TO BLOCK OR TO STOP
'-//=============================================================================================
    Cancel = PopUp
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

End Sub


Private Sub WB_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

'-//===========================================================================================
'-// SHOWS PROGRESS BAR
'-//============================================================================================
            On Error Resume Next
            PB.Max = ProgressMax
            PB.Value = Progress
            PB.Refresh
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

End Sub

Private Sub WB_StatusTextChange(ByVal Text As String)

'-//=============================================================================================
'-//DISPLAY IN STATUSBAR, THE STATUS OF THE BROWSER
'-//DOWNLOAD STATUS AND HYPERLINK FLY-OVERS
'-//=============================================================================================
      
    If Len(Text) Then
         StatusBar.Panels(1).Text = Text
    Else
        StatusBar.Panels(1).Text = WB.LocationName
    End If
  
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/

'-//============================================================================================
'-// RESIZE/ADJUST THE STATUS BAR
'-//============================================================================================
 
lwidth = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width)
        
        If lwidth > 0 Then
        
           StatusBar.Panels(1).Width = lwidth
           StatusBar.Panels(1).Visible = True
                 
        Else
        
           StatusBar.Panels(1).Visible = False
        
        End If
    '/--------------------------------------------------------------\
    '----------------------END CODING-------------------------------|
    '\--------------------------------------------------------------/
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 
    '-//========================================================================================
    '-// NAVIGATE
    '-//========================================================================================
    If cboAddress.Text <> "" Then WB.Navigate cboAddress.Text
    
    Set rsAutoComplete = New ADODB.Recordset
    rsAutoComplete.Open "SELECT * FROM TBL_AUTOCOMPLETE WHERE url='" & cboAddress.Text & "'", CN, adOpenStatic, adLockOptimistic
    
    If rsAutoComplete.EOF Then
        With rsAutoComplete
        
            .AddNew
            !URL = cboAddress
            .Update
            .Requery
        
        End With
        cboAddress.AddItem cboAddress.Text
    End If
        
End If


End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)

'-//========================================================================================================
'-// SHOW ITS NAME IN THE TITLE BAR
'-//========================================================================================================

     lblLocationname.Caption = " -   " & WB.LocationName
      cboAddress = WB.LocationURL
          
'/--------------------------------------------------------------\
'----------------------END CODING-------------------------------|
'\--------------------------------------------------------------/
End Sub

Private Sub MenuXP_Click(ByVal ID As Long)
   Dim c As Long
    
With MenuXP.MenuItems
         
        Select Case .Key(ID)
        Case "keyExit"
                
                End
                
        Case "KeyOpen"
           frmOpen.Show vbModal
        
        Case "keySave"
            On Error GoTo errs
            
           WB.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
            
        Exit Sub
        
errs:
        MsgBox "PowerSurf failed to execute!", vbCritical, "ERROR"
        
        Case "keyPageSetup"
            
            On Error GoTo errs1
            WB.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
            
  Exit Sub
        
errs1:
        MsgBox "PowerSurf failed to execute!", vbCritical, "ERROR"
        
        Case "keyPrint"
        
        On Error GoTo errs2
            WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
        
      Exit Sub
        
errs2:
        MsgBox "PowerSurf failed to execute!", vbCritical, "ERROR"
        
        Case "keyPrintPreview"
        
        On Error GoTo errs4
        
            WB.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
 Exit Sub
        
errs4:
        MsgBox "PowerSurf failed to execute!", vbCritical, "ERROR"
        
        Case "KeyProperties"
            
            WB.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
           
        Case "KeyOffline"
        
        '-//=============================================================================================
        '-// WORK OFFLINE
        '-//=============================================================================================
             MenuXP.MenuItems.Value("KeyOffline") = Not smiChecked
             
             If MenuXP.MenuItems.Value("KeyOffline") = smiChecked Then
                MenuXP.MenuItems.Value("KeyOffline") = smiChecked
                WB.Offline = True
                
            ElseIf MenuXP.MenuItems.Value("KeyOffline") = smiUnchecked Then
                
                MenuXP.MenuItems.Value("KeyOffline") = smiUnchecked
                WB.Offline = False
                        
            End If
           
        Case "KeyNewWin"
            Shell App.Path & "/PowerSurf.exe", vbNormalFocus
        
        Case "KeyContact"
        
            On Error GoTo Errhand
        
            Dim contact
            contact = Shell("C:\Program Files\Outlook Express\WAB.EXE", vbNormalFocus)
                    Exit Sub
Errhand:
        MsgBox "No installed Address Book on your computer." & vbCrLf & _
               "Or maybe installed in different location.", vbCritical

        Case "KeyNetMet"
           
           On Error GoTo ErrHandler

            Dim netmeeting
            netmeeting = Shell("C:\Program Files\NetMeeting\Conf.exe", vbNormalFocus)
            
            Exit Sub
            
ErrHandler:
            
            MsgBox "No installed Net Meeting on your computer." & vbCrLf & _
                   "Or maybe installed in different location.", vbCritical
       
            
        Case "KeyCheck"
        
            On Error GoTo err

            Shell "C:\Program Files\Outlook Express\msimn.exe", vbNormalFocus
            
            Exit Sub
err:
            MsgBox "No installed Yahoo Messenger on your computer!" & vbCrLf & _
                   "Or maybe installed in different location.", vbCritical
       

        Case "KeySend"
        
             Dim Person, Subject As String
    
    
            Person = InputBox("Enter email address", "email")
            Subject = InputBox("Enter subject for email", "subject")
            
            ShellExecute hWnd, "open", "mailto:" & Person & "?subject=" & Subject, "", "", vbNormalFocus
        
        Case "Keysynch"
        
        On Error GoTo errh
            
            Dim X, y
            X = Shell("C:\WINDOWS\SYSTEM32\mobsync.exe", vbNormalFocus)
            Exit Sub
            
errh:
            MsgBox "Items to Synchronize, Not Found on your System!", vbCritical
            
        
        Case "KeyCut"
            On Error Resume Next
                WB.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
        Case "KeyCopy"
            On Error Resume Next
                WB.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
        Case "KeyPaste"
            On Error Resume Next
                WB.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
        Case "KeyAll"
            On Error Resume Next
                WB.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
                   
        Case "KeyFind"
        
                WB.SetFocus
                SendKeys "^f"
                
        Case "KeyWinUpdate"
        
                WB.Navigate "http://windowsupdate.microsoft.com"
                
        Case "KeyYahoo"
        
                    On Error GoTo err1
            Shell "C:\Program Files\Yahoo!\Messenger\YahooMessenger.exe", vbNormalFocus
            Exit Sub
err1:
            MsgBox "No installed Yahoo Messenger on your computer!" & vbCrLf & _
                   "Or maybe installed in different location.", vbCritical

        Case "KeyMSN"
        
        
        On Error GoTo err2

        Shell ("C:\Program Files\MSN Messenger\msnmsgr.exe"), vbNormalFocus
Exit Sub
err2:
        MsgBox "No installed MSN Messenger on your computer." & vbCrLf & _
       "Or maybe installed in different location.", vbCritical

        Case "KeyIEopt"
        Shell ("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0"), vbNormalFocus

        Case "KeyInfo"
            frmProfile.List1.Selected(0) = True
            frmProfile.Show , Me
            
        Case "KeySource"
        
              frmsource.Show vbModal
              
                  
        Case "KeyLargest"
             
             MenuXP.MenuItems.Value("KeyLargest") = smiChecked
             MenuXP.MenuItems.Value("KeyLarge") = smiUnchecked
             MenuXP.MenuItems.Value("KeyMedium") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmall") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmallest") = smiUnchecked
             
             WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
             
        Case "KeyLarge"
             
             MenuXP.MenuItems.Value("KeyLargest") = smiUnchecked
             MenuXP.MenuItems.Value("KeyLarge") = smiChecked
             MenuXP.MenuItems.Value("KeyMedium") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmall") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmallest") = smiUnchecked
             
            WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
            
        Case "KeyMedium"
        
             MenuXP.MenuItems.Value("KeyLargest") = smiUnchecked
             MenuXP.MenuItems.Value("KeyLarge") = smiUnchecked
             MenuXP.MenuItems.Value("KeyMedium") = smiChecked
             MenuXP.MenuItems.Value("KeySmall") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmallest") = smiUnchecked
             
            WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
            
         Case "KeySmall"
         
             MenuXP.MenuItems.Value("KeyLargest") = smiUnchecked
             MenuXP.MenuItems.Value("KeyLarge") = smiUnchecked
             MenuXP.MenuItems.Value("KeyMedium") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmall") = smiChecked
             MenuXP.MenuItems.Value("KeySmallest") = smiUnchecked
             WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
            
         Case "KeySmallest"
         
             MenuXP.MenuItems.Value("KeyLargest") = smiUnchecked
             MenuXP.MenuItems.Value("KeyLarge") = smiUnchecked
             MenuXP.MenuItems.Value("KeyMedium") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmall") = smiUnchecked
             MenuXP.MenuItems.Value("KeySmallest") = smiChecked
             
             WB.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
            
        Case "KeyRefresh"
            On Error Resume Next
            WB.Refresh
            
        Case "KeyStop"
            On Error Resume Next
            WB.Stop
            
        Case "Keyhome"
         On Error Resume Next
            WB.GoHome
            
        Case "KeySearch"
         On Error Resume Next
            WB.GoSearch
            
        Case "KeyAdd"
        
            If MsgBox("Do You Wish To Add   [ " & cboAddress.Text & " ]    To Your Favorites List ", vbInformation + vbYesNo, " Information. ") = vbNo Then Exit Sub

            FavoriteAddress = cboAddress.Text
            Open App.Path + "/Favorites.ini" For Append As #1
            Write #1, FavoriteAddress
            Close #1
                
         Case "Keyabout"
             frmabout.Show vbModal
             
        Case "KeyOpt"
        
            '-//==================================================================================
            '-// VERIFY IF PASSWORD FOR BLOCK/LOCK APPLICATION HAS BEEN SET
            '-// IF NOT FRMSETPASSWORD WILL SHOW UP
            '-//==================================================================================

            Set rschkPass = New ADODB.Recordset
            rschkPass.Open "SELECT * FROM TBL_LOCK", CN, adOpenStatic, adLockOptimistic
            
               If rschkPass.EOF Then
                    MsgBox "Password has not yet been set. Please set new password before using this Features." & vbCrLf _
                    & "Go To --> Tools Menu --> Lock PowerSurf, to set new password." & vbCrLf & vbCrLf & vbCrLf _
                    & "                                                                                  - RJ GALLERMO", vbExclamation, "REMINDER"
                    Exit Sub
                Else
                   frmoptions.Show vbModal
                End If
                
            Set rschkPass = Nothing
                      
        Case "KeyBack"
            On Error Resume Next
            WB.GoBack
            
        Case "KeyForward"
            On Error Resume Next
            WB.GoForward
            
        Case "KeyPopblock"
        
            '-//================================================================================
            '-// TURN ON/OFF POP-UP WINDOWS
            '-//================================================================================
                      
                If PopUp = True Then
                    PopUp = False
                    MenuXP.MenuItems.Add "KeyPop", "KeyPopblock", smiCheckBox, "Allow Pop-up", , , , smiUnchecked
                    StatusBar.Panels(3) = "Enable Pop-Ups"
                     Call SaveSetting("PowerSurf", "Option", "Popup", "1")
                Else
                    PopUp = True
                    MenuXP.MenuItems.Add "KeyPop", "KeyPopblock", smiCheckBox, "Allow Pop-up", , , , smiChecked
                    StatusBar.Panels(3) = "Disable Pop-Ups"
                    Call SaveSetting("PowerSurf", "Option", "Popup", "0")
                End If

        Case "KeyWindowsEx"
            Shell "explorer.exe", vbNormalFocus
    
        Case "Keyvercheck"
                frmcheckversion.Show vbModal
                
        Case "KeyStartMenu"
           
            NavStartMenu
    
        Case "KeyPrograms"
            NavPrograms
        
        Case "KeyStartUp"
            NavStartUp
            
        Case "KeyRecent"
            NavRecent
            
        Case "KeyDesktop"
            NavDesktop
            
        Case "KeyFAvoritesShow"
            NavFavorites
            
        Case "Keyhood"
            NavNetHood
            
        Case "KeyPlayer"
          
            hist = False
            cmdClose_Click
            
            sPlay = True
            viewplayer

        Case "KeyH"
        
            '-//=================================================================================
            '-// IT USES MICROSOFT SCRIPTING RUNTIME
            '-//================================================================================
            Dim fso As New FileSystemObject
                         
            If fso.FileExists(App.Path & "/powersurf.chm") Then
            
            '-//==================================================================================
            '-// CALL HELP FILE, IF NOT EXIST THEN MSGBOX WILL BE DISPLAYED
            '-//==================================================================================
            ShellExecute hWnd, "open", App.Path & "/powersurf.chm", "", "", vbNormalFocus
   
            Else
            
                MsgBox ("PowerSurf Help file not found."), vbInformation
                Exit Sub
                
            End If
          
        Case "KeyLock"
          
            '-//==================================================================================
            '-// VERIFY IF PASSWORD FOR BLOCK/LOCK APPLICATION HAS BEEN SET
            '-// IF NOT FRMSETPASSWORD WILL SHOW UP
            '-//==================================================================================

            Set rschkPass = New ADODB.Recordset
            rschkPass.Open "SELECT * FROM TBL_LOCK", CN, adOpenStatic, adLockOptimistic
            
                If rschkPass.EOF Then
                    frmsetpassword.Show vbModal
                Else
                    frmMain.Enabled = False
                    frmlock.Show vbModal
                End If
                
            Set rschkPass = Nothing
            
          Case "KeyChangeskin"
                picskin.Visible = True
                
          Case "KeyShort"
            frmshorcut.Show vbModal
            
          Case "KeyViewFav"
                frmViewFav.Show , Me
            
            
      End Select
       
    End With
    
End Sub


