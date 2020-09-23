Attribute VB_Name = "Module7"


Public Sub GetRegistryValues1(n As Integer)

' this code on upon form activate gets the values for each
' of the options and set them...


    Dim i As Integer
    Dim value As Integer
    Dim val
    
       
    
For i = 0 To n - 1
     
  Select Case i
      
                 
        Case 0:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Accessibility")
                
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
           
        Case 1:
        
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "GeneralTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
              
              
        Case 2:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecurityTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 3:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ContentTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 4:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ConnectionsTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 5:
                                
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ProgramsTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        
        Case 6:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "AdvancedTab")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                        
        
        Case 7:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifPers")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        
        Case 8:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifSite")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 9:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecChangeSettings")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 10:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecAddSites")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 11:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 12:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest Passwords")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 13:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Connwiz Admin Lock")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 14:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Settings")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 15:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ResetWebSettings")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        
        Case 16:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpMenu")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        Case 17:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemNetscapeHelp")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        Case 18:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemSendFeedback")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        
        Case 19:
                                
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTipOfTheDay")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        
        Case 20:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTutorial")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        
        Case 21:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserClose")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        
        Case 22:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserContextMenu")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 23:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserOptions")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 24:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveAs")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 25:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFavorites")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 26:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileNew")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 27:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileOpen")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 28:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFindFiles")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        
        Case 29:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSelectDownloadDir")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 30:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoTheaterMode")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 31:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoAddressBar")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 32:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoToolBar")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 33:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoLinksBar")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 34:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoToolbarCustomize")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
                
        Case 35:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoBandCustomize")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If

        Case 36:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Use Custom Search URL")
                If value = 0 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 37:
                val = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "ShowGoButton")
                
                If val = "no" Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
           
                 
        Case 38:
                val = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NotifyDownloadComplete")
                If val = "no" Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        Case 39:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NoUpdateCheck")
                If value = 1 Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
        
        
        Case 40:
                val = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger")
                If val = "yes" Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If
                
        
        Case 41:
                val = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Ftp", "Use Web Based FTP")
                If val = "No" Then
                    Form1.List4.Selected(i) = True
                Else
                    Form1.List4.Selected(i) = False
                End If

  End Select
   
Next i


End Sub



Public Sub Registry1(n As Integer)

' coding for the apply button on registry tab
' internet explorer tab

Dim i As Integer
    
    
For i = 0 To n - 1
     
    Select Case i
      
      
        Case 0:
                ' The features available within the
                ' Internet Explorer control panel
                ' (Tools -> Internet Options) can be
                ' individually managed and disabled using
                ' this tweak. case 0 to case 15
                
                'Disable all options under Accessibility
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Accessibility", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Accessibility"
           
                End If

        
        
        Case 1:
        
                ' Remove General tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "GeneralTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "GeneralTab"
           
                End If
        
        
        Case 2:
                ' Remove Security tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecurityTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecurityTab"
           
                End If
        
        
        Case 3:
                ' Remove Content tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ContentTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ContentTab"
           
                End If
        
        
        Case 4:
                'Remove Connections tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ConnectionsTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ConnectionsTab"
           
                End If
        
        
        Case 5:
                'Remove Programs tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ProgramsTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ProgramsTab"
           
                End If
                
        
        
        Case 6:
                ' Remove Advanced tab
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "AdvancedTab", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "AdvancedTab"
           
                End If
                
                        
        
        Case 7:
                ' Prevent changing Certificate options
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifPers", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifPers"
           
                End If
        
        
        
        Case 8:
                ' Remove the Personal tab from
                ' Certificate manager
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifSite", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "CertifSite"
           
                End If
        
        
        Case 9:
                'Prevent changing Security Levels for the
                ' Internet Zone
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecChangeSettings", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecChangeSettings"
           
                End If
        
        
        Case 10:
                ' Prevent adding Sites to ANY zone
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecAddSites", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "SecAddSites"
           
                End If
        
        
        Case 11:
                ' Disable AutoComplete for forms
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest"
           
                End If
        
        
        Case 12:
                ' Prevent Prompt me to save
                ' password from being displayed
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest Passwords", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "FormSuggest Passwords"
           
                End If
        
        
        Case 13:
                ' Disable the Internet Connection Wizard
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Connwiz Admin Lock", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Connwiz Admin Lock"
           
                End If
        
        
        Case 14:
                ' Prevent any changes to Temporary
                ' Internet Files
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Settings", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Settings"
           
                End If
        
        
        Case 15:
                ' Disable the Reset web Setting button
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel"
                SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ResetWebSettings", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "ResetWebSettings"
           
                End If
        
        
        Case 16:
                ' Restrict Help Menu Items in Internet Explorer (All Versions)
                ' The menu items with the Internet Explorer
                ' "Help" menu can be individually removed or
                ' the menu disabled completely using this tweak.
                ' case 16 to case 20
                
                ' Disables the entire help menu

                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpMenu", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpMenu"
           
                End If
        
        Case 17:
                ' Remove the "For Netscape Users" menu item
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemNetscapeHelp", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemNetscapeHelp"
           
                End If
        
        Case 18:
                'Remove the "Send Feedback" menu item
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemSendFeedback", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemSendFeedback"
           
                End If
        
        
        Case 19:
                ' Removes the "Tip of the Day" menu item
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTipOfTheDay", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTipOfTheDay"
           
                End If
        
        
        Case 20:
                ' Remove the "Tour" (Tutorial) menu item
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTutorial", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoHelpItemTutorial"
           
                End If
                
        
        
        Case 21:
                ' Internet Explorer 5 Restrictions (All Versions)
                ' Microsoft Internet Explorer 5 has a range
                ' of features that can be selectively controlled
                ' by modifying the Windows registry.
                ' case 21 to case
                ' Disable the option of closing Internet Explorer.

                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserClose", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserClose"
           
                End If
        
        
        Case 22:
                ' Disable right-click context menu.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserContextMenu", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserContextMenu"
           
                End If
        
        
        Case 23:
                ' Disable the Tools / Internet Options menu.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserOptions", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserOptions"
           
                End If
        
        
        
        Case 24:
                ' Disable the ability to Save As.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveAs", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveAs"
           
                End If
        
        
        Case 25:
                ' Disable the Favorites.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFavorites", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFavorites"
           
                End If

        
        
        Case 26:
                ' Disable the File / New command.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileNew", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileNew"
           
                End If

        
        
        Case 27:
                ' Disable the File / Open command.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileOpen", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileOpen"
           
                End If

        
        
        Case 28:
                ' Disable the Find Files command.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFindFiles", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFindFiles"
           
                End If

        
        
        Case 29:
                ' Disable the option of selecting a download directory.
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSelectDownloadDir", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSelectDownloadDir"
           
                End If
        
        
        Case 30:
                ' Disable the Full Screen view option.
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoTheaterMode", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoTheaterMode"
           
                End If
        
        
        Case 31:
                ' Disable the address bar.
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoAddressBar", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoAddressBar"
           
                End If

        
        
        Case 32:
                ' Disable the tool bar.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoToolBar", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoToolBar"
           
                End If

        
        
        
        Case 33:
                ' Disable the links bar.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoLinksBar", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoLinksBar"
           
                End If



        
        
        Case 34:
                ' Disable the Ability to Customize Toolbars (All Versions)
                ' By right clicking on a toolbar you are
                ' usually given the option to Customize,
                ' which allows you to change which functions
                ' are available from the toolbar. This tweak
                ' allows you to disable that function.
                
               If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoToolbarCustomize", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoToolbarCustomize"
           
                End If
                
                
                
        Case 35:
                ' Remove the Option to Change or Hide Toolbars (All Versions)
                ' By default users are able to select which
                ' toolbars are displayed either be right
                ' clicking the toolbar itself, or by
                ' changing the options from the View menu.
                ' This tweak locks the toolbars, removing
                ' the ability to change which are displayed.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoBandCustomize", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoBandCustomize"
           
                End If


        Case 36:
                ' Disable the Custom Search Page in Internet Explorer (All Versions)
                ' This setting allows you to diable the use
                ' of the custom search page in Internet
                ' Explorer.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Use Custom Search URL", REG_DWORD, "0"

                Else
           
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Use Custom Search URL", REG_DWORD, "1"
           
                End If
        
        
        Case 37:
                ' Disable the Go Button in Internet Explorer (All Versions)
                ' This setting is used to remove the "Go"
                ' button from the Internet Explorer toolbar.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "ShowGoButton", REG_SZ, "no"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "ShowGoButton"
           
                End If

        
        
        Case 38:
                ' Disable Internet Explorer Download Notification (All Versions)
                ' This setting is used to disable download
                ' notification in Internet Explorer 5.0.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NotifyDownloadComplete", REG_SZ, "no"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NotifyDownloadComplete"
           
                End If
        
        
        Case 39:
                ' Check for Internet Explorer Updates (All Versions)
                ' Internet Explorer 5 and higher has the
                ' ability to automatically check for software
                ' updates. This tweak controls that feature.
                
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NoUpdateCheck", REG_DWORD, "1"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NoUpdateCheck"
           
                End If
                
        
        
        Case 40:
                'Control the Internet Explorer Script Debugger (All Versions)
                ' When an Internet Explorer detects an error
                ' on a page it has the ability to launch a
                ' script debugger to diagnose the problem.
                ' This setting controls the use of the Internet
                ' Explorer script debugging functions.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger", REG_SZ, "yes"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger"
           
                End If
                
                
        
        Case 41:
                ' Internet Explorer FTP Mode (All Versions)
                ' Internet Explorer has the ability to
                ' display FTP sites as if they were local
                ' folders. This tweak controls which mode
                ' IE uses for FTP.
                
                If Form1.List4.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Ftp"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Ftp", "Use Web Based FTP", REG_SZ, "No"

                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Ftp", "Use Web Based FTP"
           
                End If
        
        
        
    
    
    End Select
   
Next i

End Sub
