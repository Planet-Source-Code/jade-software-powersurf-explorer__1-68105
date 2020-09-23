Attribute VB_Name = "ModMenus"
Option Explicit
Public X As New CLS_MenuXP '-// VARIABLE FOR MENU FUNCTIONS

'-//==================================================================================================
'-// PROCEDURE FOR BUILDING MENUS
'-//==================================================================================================

Public Sub pBuildMenus()

    With frmMain.MenuXP.MenuItems
    
    'Root > File
     .Add 0, "keyfile", , "&File"
     .Add "keyFile", "keyNew", , "&New"
     .Add "keyFile", "KeyOpen", smiPicture, "Op&en...", X.pGetPicture("OPenfile"), vbCtrlMask, vbKeyO
     .Add "keyFile", "keySave", , "&Save As...", X.pGetPicture("saveas")
     .Add "keyFile", , smiSeparator
     .Add "keyFile", "keyPageSetup", , "Pa&ge Setup..."
     .Add "keyFile", "keyPrint", , "&Print...", X.pGetPicture("Printer"), vbCtrlMask, vbKeyP
     .Add "keyFile", "keyPrintPreview", , "Print Pre&view...", X.pGetPicture("PreviewPrint")
     .Add "keyFile", , smiSeparator
     .Add "KeyFile", "KeyProperties", , "P&roperties"
     .Add "KeyFile", "KeyOffline", smiCheckBox, "&Work Offline", , , , smiUnchecked
        
    'Root > File > New
     .Add "keyNew", "KeyNewWin", smiPicture, "&Window   ", , vbCtrlMask, vbKeyN
     .Add "keyNew", "KeyContact", smiPicture, "&Contact  "
     .Add "keyNew", , smiSeparator
     .Add "keyNew", "KeyNetMet", smiPicture, "&Net Meeting  "
     .Add "keyFile", , smiSeparator
     .Add "keyFile", "keyExit", , "E&xit     ", X.pGetPicture("Exit"), vbAltMask, vbKeyQ
    
    
    'Root > Edit
     .Add 0, "keyEdit", , "&Edit"
     .Add "keyEdit", "KeyCut", smiPicture, "&Cut  ", X.pGetPicture("cut"), vbCtrlMask, vbKeyX
     .Add "keyEdit", "KeyCopy", smiPicture, "Co&py  ", X.pGetPicture("copy"), vbCtrlMask, vbKeyC
     .Add "keyEdit", "KeyPaste", smiPicture, "&Paste  ", X.pGetPicture("paste"), vbCtrlMask, vbKeyV
     .Add "keyEdit", , smiSeparator
     .Add "keyEdit", "KeyAll", smiPicture, "Select &All ", X.pGetPicture("all"), vbCtrlMask, vbKeyA
     .Add "keyEdit", , smiSeparator
     .Add "keyEdit", "KeyFind", smiPicture, "&Find (on this page) ", X.pGetPicture("find"), vbCtrlMask, vbKeyF
     
    'Root > View
     .Add 0, "KeyView", , "&View"
     .Add "KeyView", "KeyGoto", smiPicture, "&Go To"
     .Add "KeyView", , smiSeparator
     .Add "KeyView", "KeyStop", smiPicture, "&Stop", , , vbKeyEscape
     .Add "KeyView", "KeyRefresh", smiPicture, "&Refresh ", X.pGetPicture("refresh"), , vbKeyF5
     .Add "KeyView", , smiSeparator
     .Add "KeyView", "KeySize", smiPicture, "&Text Size"
     .Add "KeyView", "KeySource", smiPicture, "&View Source", X.pGetPicture("ViewSource")
        
    'Root > View > Goto
    .Add "KeyGoto", "KeyBack", smiPicture, "&Back"
    .Add "KeyGoto", "KeyForward", smiPicture, "&Forward"
    .Add "KeyGoto", , smiSeparator
    .Add "KeyGoto", "Keyhome", smiPicture, "&Homepage   ", X.pGetPicture("home"), vbCtrlMask, vbKeyH
    .Add "KeyGoto", "KeySearch", smiPicture, "Search the &Web...", X.pGetPicture("search")
    
    'Root > View > Text Size
    
    .Add "KeySize", "KeyLargest", smiCheckBox, "&Largest"
    .Add "KeySize", "KeyLarge", smiCheckBox, "Lar&ge"
    .Add "KeySize", "KeyMedium", smiCheckBox, "&Medium", , , , smiChecked
    .Add "KeySize", "KeySmall", smiCheckBox, "&Small"
    .Add "KeySize", "KeySmallest", smiCheckBox, "S&mallest"
    
    
    '-//Root > Favorites
     .Add 0, "KeyFav", , "F&avorites"
     .Add "KeyFav", "KeyAdd", smiPicture, "&Add to Favorites..."
     .Add "KeyFav", "KeyViewFav", smiPicture, "&View Favorites..."
     
    '-//Root > Tools
     .Add 0, "KeyTools", , "&Tools"
     .Add "KeyTools", "KeyShortCuts", , "S&horcuts"
     .Add "KeyTools", , smiSeparator
     .Add "keyTools", "keyMail", , "&Mail"
     .Add "KeyTools", "Keysynch", smiPicture, "&Synchronize"
     .Add "KeyTools", "KeyWinUpdate", smiPicture, "&Windows Update", X.pGetPicture("WindowsUpdate")
     .Add "KeyTools", , smiSeparator
     .Add "KeyTools", "KeyYahoo", smiPicture, "&Yahoo! Messenger", X.pGetPicture("ym")
     .Add "KeyTools", "KeyMSN", smiPicture, "MS&N! Messenger", X.pGetPicture("msn")
     .Add "KeyTools", , smiSeparator
     .Add "KeyTools", "KeyPop", , "&Pop-Up Blocker"
     .Add "KeyTools", , smiSeparator
     .Add "KeyTools", "KeyOpt", smiPicture, "&Options...     ", X.pGetPicture("options"), , vbKeyF7
     .Add "KeyTools", "KeyIEopt", smiPicture, "&Internet Options...      ", , , vbKeyF8
     .Add "KeyTools", "KeyInfo", smiPicture, "My Personal In&formation...", X.pGetPicture("Personal")
     .Add "KeyTools", , smiSeparator
     .Add "KeyTools", "KeyLock", smiPicture, "&Lock PowerSurf", X.pGetPicture("LockApp")
     .Add "KeyTools", "KeyPlayer", smiPicture, "PowerSurf WinMedia &Fx", X.pGetPicture("player")
     .Add "KeyTools", "KeyChangeskin", smiPicture, "&Change Skin", X.pGetPicture("skin"), vbCtrlMask, vbKeyS
          
     'Root > Tools > Shorcuts
     .Add "KeyShortcuts", "KeyWindowsEx", smiPicture, "Windows &Explorer", X.pGetPicture("explorer"), vbCtrlMask, vbKeyE
     .Add "KeyShortcuts", , smiSeparator
     .Add "KeyShortcuts", "KeyPrograms", smiPicture, "&Programs", X.pGetPicture("programs")
     .Add "KeyShortcuts", "KeyStartMenu", smiPicture, "&Start Menu Folder", X.pGetPicture("startmenu")
     .Add "KeyShortcuts", "KeyStartUp", smiPicture, "Start-&Up Folder", X.pGetPicture("startup")
     .Add "KeyShortcuts", , smiSeparator
     .Add "KeyShortcuts", "KeyRecent", smiPicture, "&Recent", , vbCtrlMask, vbKeyR
     .Add "KeyShortcuts", "KeyDesktop", smiPicture, "&Desktop", X.pGetPicture("desktop")
     .Add "KeyShortcuts", "KeyFAvoritesShow", smiPicture, "&Favorites", X.pGetPicture("Fav")
     .Add "KeyShortcuts", , smiSeparator
     .Add "KeyShortcuts", "Keyhood", smiPicture, "&Network Neighborhood", X.pGetPicture("net")

     '-//Root > Tools > Mail
     .Add "keyMail", "KeyCheck", smiPicture, "&Check Mail"
     .Add "keyMail", "KeySend", smiPicture, "&Send Mail"
      
       '-//Root > Tools > Pop-Up
     .Add "KeyPop", "KeyPopblock", smiCheckBox, "Allow Pop-up"
        
             
     '-//Root > Help
     .Add 0, "keyHelp", , "&Help"
     .Add "KeyHelp", "KeyH", smiPicture, "&Contents", X.pGetPicture("Help"), , vbKeyF1
     .Add "KeyHelp", "Keyvercheck", smiPicture, "Check New &Version...", X.pGetPicture("CheckVer")
     .Add "KeyHelp", "KeyShort", smiPicture, "&Keyboard Shortcut", X.pGetPicture("ShortKey"), vbCtrlMask, vbKeyF2
     .Add "keyHelp", , smiSeparator
     .Add "KeyHelp", "Keyabout", smiPicture, "&About PowerSurf Explorer", X.pGetPicture("about")
     
    End With
End Sub
