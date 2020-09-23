Attribute VB_Name = "ModSkin"
Option Explicit
'-//==========================================================================================
'-// DO NOT TRY TO MODIFY THIS PART
'-//==========================================================================================

Public Sub DefaultSkin()
        
        
        frmMain.MenuXP.BackColor = &H6B553C
        frmMain.MenuXP.FontForeColor = &HFFFFFF
        frmMain.MenuXP.FontBackColor = &H847860
        frmMain.pictop.Height = 460
        frmMain.lblLocationname.Top = 110
        frmMain.Label1.Top = 110
        frmMain.pictop.BackColor = &H6B553C
        frmMain.MenuXP.Top = 465
        frmMain.CoolBar1.Top = 830
        frmMain.TBTOP.Top = 855
        frmMain.Image1.Left = 400
        frmMain.pictop.Picture = LoadPicture(App.Path & "/Skins/Default/top.skn")
        frmMain.imgmin.Picture = LoadPicture(App.Path & "/Skins/Default/MinimizeUp.skn")
        frmMain.cmglose.Picture = LoadPicture(App.Path & "/Skins/Default/CloseUp.skn")
        frmMain.BackColor = &H6B553C
        frmMain.Image1.Top = 120
        frmMain.cmglose.Top = 110
        frmMain.imgmin.Top = 110
        frmMain.imgmin.Left = 14640
        frmMain.cmglose.Left = 14955
        frmshorcut.BackColor = &H6B553C
        frmMain.lblLocationname.Left = 2760
        frmViewFav.BackColor = &H6B553C
        frmsource.BackColor = &H6B553C
        
        frmsetpassword.BackColor = &H6B553C
        frmsearch.BackColor = &H6B553C
        frmsetpassword.Frame1.BackColor = &H6B553C
        
        frmsearch.Frame1.BackColor = &H6B553C
        frmsearch.Frame2.BackColor = &H6B553C
        frmsearch.Label2.BackColor = &H6B553C
        
            Dim color As Integer
        
            For color = 0 To 14
            frmsearch.Option1(color).BackColor = &H6B553C
            Next
        
        frmProfile.BackColor = &H6B553C
        frmoptions.BackColor = &H6B553C
        frmoptions.frmePrivacy.BackColor = &H6B553C
        frmoptions.frmstartup.BackColor = &H6B553C
        frmoptions.frmblocked.BackColor = &H6B553C
        frmoptions.chkCookies.BackColor = &H6B553C
        frmoptions.chkurl.BackColor = &H6B553C
        frmoptions.ChkFav.BackColor = &H6B553C
        frmoptions.OptBlank.BackColor = &H6B553C
        frmoptions.optDefine.BackColor = &H6B553C
        frmoptions.Optgoogle.BackColor = &H6B553C
        frmoptions.Opthome.BackColor = &H6B553C
        frmoptions.OptLast.BackColor = &H6B553C
        frmoptions.OptYahoo.BackColor = &H6B553C
      
        frmOpen.BackColor = &H6B553C
        frmlock.BackColor = &H6B553C
        frmcheckversion.BackColor = &H6B553C
        frmabout.BackColor = &H6B553C
        frmabout.Frame2.BackColor = &H6B553C
        frmabout.Frame7.BackColor = &H6B553C
        frmabout.Frame7.ForeColor = vbBlue
        frmabout.Label9.ForeColor = vbBlue
        frmabout.Label13.ForeColor = vbBlue
        
        Call SaveSetting("PowerSurf", "SKIN", "Default", 1)
        Call SaveSetting("PowerSurf", "SKIN", "Wooden", 0)
        Call SaveSetting("PowerSurf", "SKIN", "Silver", 0)
        
End Sub

Public Sub WoodenSkin()

        frmMain.MenuXP.BackColor = &H5E5E5E
        frmMain.MenuXP.FontForeColor = &HFFFFFF
        frmMain.MenuXP.FontBackColor = &H8000000C
        frmMain.pictop.Height = 460
        frmMain.MenuXP.Top = 465
        frmMain.CoolBar1.Top = 830
        frmMain.TBTOP.Top = 855
        frmMain.Image1.Top = 110
        frmMain.lblLocationname.Top = 120
        frmMain.Label1.Top = 120
        frmMain.pictop.BackColor = &H5E5E5E
        frmMain.pictop.Picture = LoadPicture(App.Path & "/Skins/Skin1/top.skn")
        frmMain.imgmin.Picture = LoadPicture(App.Path & "/Skins/Skin1/MinimizeUp.skn")
        frmMain.cmglose.Picture = LoadPicture(App.Path & "/Skins/Skin1/CloseUp.skn")
        frmMain.BackColor = &H5E5E5E
        frmshorcut.BackColor = &H5E5E5E
        frmViewFav.BackColor = &H5E5E5E
        frmMain.cmglose.Top = 130
        frmMain.imgmin.Top = 130
        frmMain.imgmin.Left = 14640
        frmMain.cmglose.Left = 14955
        
        frmsource.BackColor = &H5E5E5E
        
        frmsetpassword.BackColor = &H5E5E5E
        frmsearch.BackColor = &H5E5E5E
        frmsetpassword.Frame1.BackColor = &H5E5E5E
        frmMain.lblLocationname.Left = 5480
        
        frmsearch.Frame1.BackColor = &H5E5E5E
        frmsearch.Frame2.BackColor = &H5E5E5E
        frmsearch.Label2.BackColor = &H6B553C
        
            Dim color1 As Integer
        
            For color1 = 0 To 14
            frmsearch.Option1(color).BackColor = &H5E5E5E
            Next
        
        frmProfile.BackColor = &H5E5E5E
        frmoptions.BackColor = &H5E5E5E
        frmoptions.frmePrivacy.BackColor = &H5E5E5E
        frmoptions.frmstartup.BackColor = &H5E5E5E
        frmoptions.frmblocked.BackColor = &H5E5E5E
        frmoptions.chkCookies.BackColor = &H5E5E5E
        frmoptions.chkurl.BackColor = &H5E5E5E
        frmoptions.ChkFav.BackColor = &H5E5E5E
        frmoptions.OptBlank.BackColor = &H5E5E5E
        frmoptions.optDefine.BackColor = &H5E5E5E
        frmoptions.Optgoogle.BackColor = &H5E5E5E
        frmoptions.Opthome.BackColor = &H5E5E5E
        frmoptions.OptLast.BackColor = &H5E5E5E
        frmoptions.OptYahoo.BackColor = &H5E5E5E
    
        frmOpen.BackColor = &H5E5E5E
        frmlock.BackColor = &H5E5E5E
        frmcheckversion.BackColor = &H5E5E5E
        frmabout.BackColor = &H5E5E5E
        frmabout.Frame2.BackColor = &H5E5E5E
        frmabout.Frame7.BackColor = &H5E5E5E
        frmabout.Frame7.ForeColor = vbBlue
        frmabout.Label9.ForeColor = vbWhite
        frmabout.Label13.ForeColor = vbWhite
        
        Call SaveSetting("PowerSurf", "SKIN", "Default", 0)
        Call SaveSetting("PowerSurf", "SKIN", "Wooden", 1)
        Call SaveSetting("PowerSurf", "SKIN", "Silver", 0)
  
 
End Sub

Public Sub SilverGraySkin()

        frmMain.MenuXP.BackColor = &H979894
        frmMain.MenuXP.FontForeColor = &HFFFFFF
        frmMain.MenuXP.FontBackColor = &H979894
        frmMain.pictop.Height = 460
        frmMain.MenuXP.Top = 465
        frmMain.CoolBar1.Top = 830
        frmMain.TBTOP.Top = 855
        frmMain.pictop.BackColor = &HA1A29E
        frmMain.pictop.Picture = LoadPicture(App.Path & "/Skins/Skin3/top.skn")
        frmMain.imgmin.Picture = LoadPicture(App.Path & "/Skins/Skin3/MinimizeUp.skn")
        frmMain.cmglose.Picture = LoadPicture(App.Path & "/Skins/Skin3/CloseUp.skn")
        frmMain.BackColor = &HA1A29E
        
        frmMain.cmglose.Top = 60
        frmMain.imgmin.Top = 60
        frmMain.imgmin.Left = 14640
        frmMain.cmglose.Left = 14955
        frmMain.Label1.Top = 50
        frmMain.lblLocationname.Top = 50
        frmMain.lblLocationname.Left = 2760
        frmMain.Image1.Top = 50
        frmMain.Image1.Left = 220
        
        frmsource.BackColor = &HA1A29E
        frmViewFav.BackColor = &HA1A29E
        frmsetpassword.BackColor = &HA1A29E
        frmsearch.BackColor = &HA1A29E
        frmsetpassword.Frame1.BackColor = &HA1A29E
        frmshorcut.BackColor = &HA1A29E
        frmsearch.Frame1.BackColor = &HA1A29E
        frmsearch.Frame2.BackColor = &HA1A29E
        frmsearch.Label2.BackColor = &HA1A29E
        
            Dim color2 As Integer
        
            For color2 = 0 To 14
            frmsearch.Option1(color2).BackColor = &HA1A29E
            Next
        
        frmProfile.BackColor = &HA1A29E
        frmoptions.BackColor = &HA1A29E
        frmoptions.frmePrivacy.BackColor = &HA1A29E
        frmoptions.frmstartup.BackColor = &HA1A29E
        frmoptions.frmblocked.BackColor = &HA1A29E
        frmoptions.chkCookies.BackColor = &HA1A29E
        frmoptions.chkurl.BackColor = &HA1A29E
         frmoptions.ChkFav.BackColor = &HA1A29E
        frmoptions.OptBlank.BackColor = &HA1A29E
        frmoptions.optDefine.BackColor = &HA1A29E
        frmoptions.Optgoogle.BackColor = &HA1A29E
        frmoptions.Opthome.BackColor = &HA1A29E
        frmoptions.OptLast.BackColor = &HA1A29E
        frmoptions.OptYahoo.BackColor = &HA1A29E
      
        frmOpen.BackColor = &HA1A29E
        frmlock.BackColor = &HA1A29E
        frmcheckversion.BackColor = &HA1A29E
        frmabout.BackColor = &HA1A29E
        frmabout.Frame2.BackColor = &HA1A29E
        frmabout.Frame7.BackColor = &HA1A29E
        frmabout.Frame7.ForeColor = vbBlack
        frmabout.Label9.ForeColor = vbBlack
        frmabout.Label13.ForeColor = vbBlack
        
        Call SaveSetting("PowerSurf", "SKIN", "Default", 0)
        Call SaveSetting("PowerSurf", "SKIN", "Wooden", 0)
        Call SaveSetting("PowerSurf", "SKIN", "Silver", 1)
        
End Sub

Public Sub SelectSkin(skin As Boolean)

    frmMain.MousePointer = vbHourglass

    If frmMain.List1 = "Default" Then DefaultSkin
    If frmMain.List1 = "WoodenBrown" Then WoodenSkin
    If frmMain.List1 = "SilverGray" Then SilverGraySkin
    
    frmMain.MousePointer = vbDefault
    frmMain.picskin.Visible = False
    
End Sub
