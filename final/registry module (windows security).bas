Attribute VB_Name = "Module6"


Public Sub GetRegistryValues(n As Integer)

' this code on upon form activate gets the values for each
' of the options and set them...


    Dim i As Integer
    Dim value As Integer
       
    
    For i = 0 To n - 1
     
     Select Case i
      
      Case 0:
             
             value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispAppearancePage")
             
             If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
             
                    
                    
      Case 1:
             
             value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispBackgroundPage")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
               
            
      Case 2:
                
               value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispScrSavPage")
               If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     Case 3:
            
            
              value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispSettingsPage")
              If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
             
             
     Case 4:
              
              value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharingControl")
              If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
              
             
     Case 5:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupIDPage")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 6:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupSecurityPage")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
             
             
     Case 7:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinterTabs")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 8:
            
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
               
               
     Case 9:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter")
          If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 10:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDevMgrPage")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 11:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoConfigPage")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 12:
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoFileSysPage")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
               
     
     Case 13:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoVirtMemPage")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 14:
            
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoPwdPage")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 15:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoAdminPage")
           If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 16:
          
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoProfilePage")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
               
     
     Case 17:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 18:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 19:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 20:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinters")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
              
              
     Case 21:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoSecCPL")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 22:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
                        
     
     Case 23:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "NoRealMode")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 24:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Network\Logon", "MustBeValidated")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 25:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "HideSharePwds")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
                   
     
     Case 26:
                 
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharing")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoPrintSharing")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 27:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Network\Logon", "NoDomainPwdCaching")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 28:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoEntireNetwork")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 29:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoWorkgroupContents")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 30:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
     
     Case 31:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools")
            If value = 1 Then
               Form1.List2.Selected(i) = True
             Else
               Form1.List2.Selected(i) = False
             End If
            
            
     Case 32:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Network\Logon", "UserProfiles")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
            
     
     Case 33:
                        
            Dim value1 As Variant
            value1 = QueryValue(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "NoLogon")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
            
     
     Case 34:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ForceActiveDesktopOn")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
                 
     
         
     
     Case 35:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 36:
             value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
     
     Case 37:
     
           value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate")
           If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
           
     
     Case 38:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "AlphanumPwds")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
            
     Case 39:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "DisablePwdCaching")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 40:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoDialIn")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 41:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 42:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 43:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSaveSettings")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
               
     
     Case 44:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktopChanges")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
               
     
     Case 45:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory")
           If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
               

     Case 46:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ClearRecentDocsOnExit")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 47:
     
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop")
           If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 48:
     
                         
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallpaper")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoComponents")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoAddingComponents")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoDeletingComponents")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoEditingComponents")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoCloseDragDropBands")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
             value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoMovingBands")
             If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
             
              
     
     Case 49:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetActiveDesktop")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
            
     Case 50:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoChangeStartMenu")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 51:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 52:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop")
            If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
            
     
     Case 53:
     
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives")
            If value > 0 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
              
     Case 54:
     
        value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")
        If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
        
 
 Case 55:
 
        value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood")
        If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
        
        
 Case 56:
 
        value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun")
        If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
        
 
 Case 57:
 
        value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar")
        If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If
        
        
 Case 58:
 
        value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoCommonGroups")
        If value = 1 Then
               Form1.List2.Selected(i) = True
            Else
               Form1.List2.Selected(i) = False
            End If

   End Select
   
Next i










End Sub






Public Sub Registry(n As Integer)


' coding for the apply button on registry tab
' windows security tab

    Dim i As Integer
    
    
    For i = 0 To n - 1
     
     Select Case i
      
      Case 0:
             
             ' hide the display appearence tab
             ' When enabled this setting hides the
             ' display settings appearance page.
    
             If Form1.List2.Selected(i) = True Then
               
               CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system"
               SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispAppearancePage", REG_DWORD, "1"
    
             Else
             
               Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispAppearancePage"
               
             End If
             
             
             
             
             
      Case 1:
               ' hide the display background page
               ' This option hides the background page stopping
               ' users from changing any background display
               ' settings.
               
    
               If Form1.List2.Selected(i) = True Then
               
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispBackgroundPage", REG_DWORD, "1"
    
               Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispBackgroundPage"
               
             End If
             
             
             
             
             
      Case 2:
                ' hide the screen saver setting tab
                ' This option hides the screen saver page from
                ' the display settings control, which stops
                ' users having access to change screen
                ' saver settings.
    
                If Form1.List2.Selected(i) = True Then
               
                  CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system"
                  SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispScrSavPage", REG_DWORD, "1"
    
             Else
               
               Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispScrSavPage"
               
             End If
             
             
             
                
      
     Case 3:
            
            ' hide the display setting tabs
            ' This option hides the Settings
            ' page from the display properties control.
            
            If Form1.List2.Selected(i) = True Then
               
              CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system"
              SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispSettingsPage", REG_DWORD, "1"
    
            Else
               
              Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "NoDispSettingsPage"
               
             End If
             
             
      
                
                
     Case 4:
              ' hide the file and printer sharing controls
              ' Enabling this options hides the file and
              ' printer sharing controls, stopping users from
              ' disabling or creating new file or printer shares.
    
              If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharingControl", REG_DWORD, "1"
    
              Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharingControl"
               
             End If
             
             
             
                
     Case 5:
     
            ' hide network identification page
            ' The Network Identification page include
            ' options to set the Computer Name, Workgroup
            ' and Description, enabling this option
            ' disables access to the Network ID page.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupIDPage", REG_DWORD, "1"
    
              Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupIDPage"
               
             End If
             
             
     
     
     Case 6:
     
            ' The Access Control Page, defines whether
            ' the computer support
            ' User-Level access or Share-Level access,
            ' enabling this option removes access to the
            ' Access Control Page.
    
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupSecurityPage", REG_DWORD, "1"
    
              Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupSecurityPage"
               
             End If
             
             
             
             
             
             
             
     Case 7:
     
            ' This option hides the printer details and
            ' general printer information pages. Once enabled
            ' this option stops users from changing specific
            ' printer settings.
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinterTabs", REG_DWORD, "1"
    
              Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinterTabs"
               
             End If
             
             
             
    
            
     Case 8:
            
            ' Printers can be deleted simply by any user
            ' pressing the delete key, enabling this setting
            ' stops users from being able to delete printers.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter"
               
            End If
            
            
            
            
            
     Case 9:
     
            ' Any user can add a new printer their system,
            ' this option once enabled disables the addition
            ' of new printers to the computer.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter"
               
            End If
            
            
     
     Case 10:
     
            ' Hide the Device Manager Page (Windows 9x/Me)
            ' This setting controls whether the Device Manager,
            ' under Control Panel / System is visible.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDevMgrPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDevMgrPage"
               
            End If
            
     
     Case 11:
     
            ' Hide the Hardware Profiles Page (Windows 9x/Me)
            ' This setting when enabled hides the Hardware
            ' Profiles page from the System icon on the Control
            ' Panel.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoConfigPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoConfigPage"
               
            End If
    
     
     Case 12:
            ' Hide the File System Button (Windows 9x/Me)
            ' This option hides the File System button from
            ' the System icon on the Control Panel.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoFileSysPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoFileSysPage"
               
            End If
            
    
     
     Case 13:
     
            '   Hide the Virtual Memory Button (Windows 9x/Me)
            '   This option hides the Virtual Memory button
            '   from the System icon on the Control Panel.
    
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoVirtMemPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoVirtMemPage"
               
            End If
            
            
     
     Case 14:
            
            ' Hide the Change Passwords Page (Windows 9x/Me)
            ' When this setting is enabled, users are no longer
            ' able to access the Change Passwords page.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoPwdPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoPwdPage"
               
            End If
            
            
     
     Case 15:
     
            ' Hide the Remote Administration Page (Windows 9x/Me)
            ' Enabling this function stops users from being
            ' able to change the remote administration settings
            ' for the computer.
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoAdminPage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoAdminPage"
               
            End If
            
            
            
     
     Case 16:
     
     
            ' Hide the User Profiles Page (Windows 9x/Me)
            ' The user profile page controls whether all
            ' users share or have separate user profiles,
            ' access to this page can be disabled by enabling
            ' this setting.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoProfilePage", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoProfilePage"
               
            End If
     
     
     
     Case 17:
     
            ' Hide Control Panel on Start Menu (All Versions) Popular
            ' This setting allows you to hide the Control
            ' Panel options from the Start Menu.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
               
            End If
            
            
     
     Case 18:
     
            ' Deny Access to the Display Settings (All Versions)
            ' This option disables the display settings control
            ' panel icon, and stops users from accessing any
            ' display settings.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL"
               
            End If
     
     
     
     Case 19:
     
            ' Disable Network Control Panel (Windows 9x/Me)
            ' Enabling this option disables access to the Network Control Panel icon.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup"
               
            End If
     
     
     
     Case 20:
     
            ' Disable Printers Control Panel Icon (Windows 9x/Me)
            ' This option disables access to the Printers icon
            ' in control panel, therefore stopping users from
            ' changing Printer settings.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinters", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinters"
               
            End If
     
     
     Case 21:
     
            ' Restrict Access to the Passwords Control Panel (Windows 9x/Me)
            ' This options disables access to the Passwords icon
            ' on the control panel, therefore stopping users
            ' from changing security related settings.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoSecCPL", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoSecCPL"
               
            End If
     
     Case 22:
     
            ' Disable the MS-DOS Command Prompt (All Versions)
            ' This setting allows you to disable the use of
            ' the MS-DOS command prompt in Windows.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled"
               
            End If
     
     Case 23:
     
            ' Disable Single Mode MS-DOS Applications in Windows
            ' This setting allows you to disable the use of
            ' real mode DOS applications from within the
            ' Windows shell.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "NoRealMode", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "NoRealMode"
               
            End If
     
    
     
     Case 24:
     
            ' Require Validation by Network for Windows Access (Windows 9x/Me)
            ' By default Windows 9x doesn't require a valid
            ' network username and password combination for a
            ' user to bypass the logon and gain access to the
            ' local machine. This functionality can be changed
            ' to require validation by the network before
            ' allowing access.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Network\Logon"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Network\Logon", "MustBeValidated", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Network\Logon", "MustBeValidated"
               
            End If
     
    
    
     
     Case 25:
     
            ' Hide Share Passwords with Asterisks (All Versions)
            ' This setting controls whether the password typed
            ' when accessing a file share is shown in clear
            ' text or as asterisks
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "HideSharePwds", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "HideSharePwds"
               
            End If
                   
     
     Case 26:
     
            ' Disable File and Printer Sharing (All Versions)
            ' When file and printer sharing is installed it
            ' allows users to make services available to other
            ' users on a network, this functionality can be
            ' disabled by changing this setting
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharing", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoPrintSharing", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharing"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoPrintSharing"
                 
               
            End If
     
     
     
     Case 27:
     
            ' Disable Caching of Domain Password (All Versions)
            ' Enabling this setting disables the caching of
            ' the domain passwords, and therefore passwords
            ' are required to be re-entered to access any
            ' additional domain resources.
     
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Network\Logon"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Network\Logon", "NoDomainPwdCaching", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Network\Logon", "NoDomainPwdCaching"
               
            End If
     
     
     
     Case 28:
     
            ' Remove Entire Network from Network Neighborhood (All Versions)
            ' Entire Network is an option under Network
            ' Neighborhood that allows users to see all the
            ' Workgroups and Domains on the network.
            ' Entire Network can be disabled, so users are
            ' confined to their own Workgroup or Domain.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoEntireNetwork", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoEntireNetwork"
               
            End If
     
    
            
     Case 29:
     
            ' Hide Workgroup Content from Network Neighborhood (All Versions)
            ' Enabling this option hides all Workgroup contents
            ' from being displayed in Network Neighborhood.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoWorkgroupContents", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoWorkgroupContents"
               
            End If
     
     
     Case 30:
     
            ' Remove 'Log Off Username' from the Start Menu (All Versions)
            ' This tweak allows you to remove the
            ' 'Log Off Username' option from the Start menu.
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff"
               
            End If
    
     
     Case 31:
            ' Disable Registry Editing Tools (All Versions) Popular
            ' This setting disables the ability to run the
            ' registry editing tools Regedit.exe or
            ' Regedt32.exe interactively.
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
               
            End If
    
     
     Case 32:
     
            ' Disable User Profiles (Windows 9x/Me)
            ' This setting can be used to disable the use
            ' of user profiles.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Network\Logon"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Network\Logon", "UserProfiles", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Network\Logon", "UserProfiles"
               
            End If
     
     
     Case 33:
     
            ' Force Users to Logon to Windows (Windows 9x/Me) Popular
            ' Usually users can simply press 'Cancel' at the
            ' Windows logon box to bypass the login process
            ' and gain access to the local computer.
            ' This tweak will logout the user if the
            ' authentication fails or the user clicks Cancel.
    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run"
                 SetKeyValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "NoLogon", REG_SZ, "RUNDLL32 shell32,SHExitWindowsEx 0"
    
             Else
               
                 Delete HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "NoLogon"
               
            End If
            
     
     
     Case 34:
     
            ' Force the Use of Active Desktop (All Versions)
            ' The user is normally given the option of
            ' disabling Active Desktop through the display
            ' properties. This tweak removes the ability to
            ' disable Active Desktop.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ForceActiveDesktopOn", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ForceActiveDesktopOn"
               
            End If
    
     
     
    
     Case 35:
     
            ' Disable Folder Options Menu
            ' This tweak allows you to hide the Folder Options
            ' function from the folder Tools menu. Allowing you
            ' to restrict access to numerous advanced folder
            ' features.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"
               
            End If
     
    
     
     Case 36:
             ' Disable Windows Hotkeys (All Versions)
             ' This tweak disables the use of Windows hotkeys.
             
             If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys"
               
            End If
     
     
     
     
     Case 37:
     
            ' Restrict Access to the Windows Update Feature (All Versions)
            ' Windows 98 and later Windows versions contain a
            ' feature known as Windows Update, which allows
            ' users to update Windows specific software.
            ' With this modification you can remove access to
            ' this feature from the settings sub-menu.
     
     
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate"
               
            End If
     
     
     Case 38:
     
            ' Require Alphanumeric Windows Password (All Versions)
            ' Windows by default will accept anything as a
            ' password, including nothing. This setting controls
            ' whether Windows will require a alphanumeric
            ' password, i.e. a password made from a combination
            ' of alpha (A, B, C...) and numeric (1, 2 ,3 ...)
            ' characters.
     
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "AlphanumPwds", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "AlphanumPwds"
               
            End If
            
            
            
     Case 39:
     
            ' Disabled Password Caching (All Versions)
            ' Normally Windows caches a copy of the users
            ' password on the local system to allow for
            ' additional automation, this leads to a possible
            ' security threat on some systems. Disabling
            ' caching means the users passwords are not cached
            ' locally. This setting also removes the second
            ' Windows password screen and also remove the
            ' possibility of networks passwords to get out of
            ' sync.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "DisablePwdCaching", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "DisablePwdCaching"
               
            End If
     
     
     Case 40:
     
            ' Disable Dial-In Access (All Versions)
            ' It's possible for users to setup a modem on a
            ' Windows machine, and by using Dial-up Networking
            ' allow callers to connect to the internal network.
            ' Especially in a corporate environment this can
            ' cause a major security risk.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoDialIn", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoDialIn"
               
            End If
     
     
     Case 41:
     
            ' Remove Folders from Settings on the Start Menu (All Versions)
            ' Removes the Control Panel and Printers folders
            ' from the Settings menu. Note: Removing the Taskbar,
            ' Control Panel, and Printer folders causes the
            ' Settings menu to be removed completely.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders"
               
            End If
     
     
     Case 42:
     
            ' Remove the Documents Folder from the Start Menu (All Versions)
            ' This setting can be used to remove the recent
            ' Documents folder from the Start Menu.
    
                    
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu"
               
            End If
            
     
     
     Case 43:
     
            ' Don't Save Settings at Exit (All Versions)
            ' Normally when Windows exits it saves the desktop
            ' configuration, including icon location, appearance
            ' etc. This setting disables these changes from
            ' being saved, this is useful in both a secure
            ' environment and when you don't want people to
            ' change the appearance of your desktop once you
            ' have it setup the way you like it.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSaveSettings", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSaveSettings"
               
            End If
    
     
     
     Case 44:
     
            ' Restrict Changes to Active Desktop Settings (All Versions)
            ' This tweak allows you to have Active Desktop
            ' enabled, but to restrict any changes to the
            ' settings.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktopChanges", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktopChanges"
               
            End If
     
     
     Case 45:
     
            ' Don't Add Recent Files to Documents on the Start Menu (All Versions)
            ' Normally when you open or access a document or
            ' file it is added to the list of recent documents
            ' on the Start Menu. This tweak will stop files from
            ' being added to the list.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory"
               
            End If
    
     
     
     Case 46:
     
            ' Clear Recent Documents When Windows Exits (All Versions)
            ' This tweak will clear the list of recent documents
            ' on the Start Menu when Windows exits.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ClearRecentDocsOnExit", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ClearRecentDocsOnExit"
               
            End If
    
    
     
     
     Case 47:
     
            ' Disable Active Desktop (All Versions)
            ' This tweak will disable the use of the Active
            ' Desktop feature.  (wide base -> local machine,
            ' else current user...)
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop"
               
            End If
            
    
     
     
     Case 48:
     
            ' Active Desktop Restrictions (All Versions)
            ' Features of the Windows Active Desktop can be
            ' selectively controlled by modifying options in the
            ' Windows registry. Following the instructions in
            ' this tweak.
            
            ' this can be local macine or current user. also
            ' it has various subparts which i included as
            ' combined. they are as follows...
            
            ' NoChangingWallpaper - Disable the ability to change wallpapers.
            ' NoComponents - Disable components.
            ' NoAddingComponents - Disable the ability to add components.
            ' NoDeletingComponents - Disable the ability to delete components.
            ' NoEditingComponents - Disable the ability to edit components.
            ' NoCloseDragDropBands
            ' NoMovingBands
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallpaper", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoComponents", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoAddingComponents", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoDeletingComponents", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoEditingComponents", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoCloseDragDropBands", REG_DWORD, "1"
                 SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoMovingBands", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallpaper"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoComponents"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoAddingComponents"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoDeletingComponents"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoEditingComponents"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoCloseDragDropBands"
                 Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoMovingBands"
               
            End If
    
    
    
     
     
     Case 49:
     
            ' Remove Active Desktop Options from the Settings Menu (All Versions)
            ' The tweak will remove the Active Desktop options
            ' from Settings on the Start Menu.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetActiveDesktop", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetActiveDesktop"
               
            End If
            
            
            
     Case 50:
     
            ' Disable the Ability to Modify the Start Menu (Windows 98)
            ' Normally users are able to right-click on the
            ' Start Menu and modify it using the context menu.
            ' This option give you the ability to disable this
            ' function.
            
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoChangeStartMenu", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoChangeStartMenu"
               
            End If
     
     
     
     Case 51:
     
            ' Disable the Shut Down Command (All Versions)
            ' This option allows you to stop users from being
            ' able to shutdown the computer by disabling the
            ' shut down command.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
               
            End If
            
            
     
     Case 52:
     
            ' Hide All Items on the Desktop (All Versions)
            ' Enabling this options hides all the items and
            ' programs on the Windows desktop.
            
            If Form1.List2.Selected(i) = True Then
                
                 CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", REG_DWORD, "1"
    
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop"
               
            End If
    
     
     
     Case 53:
     
            'Hide Drives in My Computer (All Versions)
            ' This setting controls which drives are visible
            ' in 'My Computer', it is possible to hide all
            ' drives or just selected ones.
            
            If Form1.List2.Selected(i) = True Then
                
               Form1.Hide
               Form2.Show
               
             Else
               
                 Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives"
               
            End If
    
     
     
     Case 54:
     
        ' Remove the Find Command From the Start Menu (All Versions)
        ' When enabled this setting removes the 'Find'
        ' command from the Start Menu.
        
        If Form1.List2.Selected(i) = True Then
            
             CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
             SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", REG_DWORD, "1"

         Else
           
             Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
           
        End If
 
 
 Case 55:
 
        ' Hide the Network Neighborhood Icon (All Versions)
        ' The Network Neighborhood icon is shown on the
        ' Windows desktop whenever Windows networking is
        ' installed, by enabling this setting the icon will
        ' be hidden.
        
        ' In addition this disables UNC capability from
        ' within the Explorer interface, including the Start
        ' menu's Run command, UNC paths configured by the
        ' administrator in Policies for shared folders,
        ' desktop icons, the Start command, and so forth.
        ' This does not impair the functionality of the
        ' command line Net.exe command.
        
        If Form1.List2.Selected(i) = True Then
            
             CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
             SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood", REG_DWORD, "1"

         Else
           
             Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood"
           
        End If
 

 
 
 Case 56:
 
        ' Remove the Run Command from the Start Menu (All Versions)
        ' Removes the user's ability to start applications
        ' or processes from the Start menu by removing the
        ' option completely.
        ' Note: If the user still has access to the MS-DOS
        ' prompt icon or command prompt, the user can start
        ' unauthorized applications.
        
        If Form1.List2.Selected(i) = True Then
            
             CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
             SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", REG_DWORD, "1"

         Else
           
             Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
           
        End If

 
 
 Case 57:
 
        ' Remove the Taskbar from Settings on the Start Menu (All Versions)
        ' Enabling this option removes the Taskbar option
        ' from Settings on the Start Menu, therefore stopping
        ' users from changing the taskbar properties.
        ' Note: Removing the Taskbar, Control Panel, and
        ' Printer folders causes the Settings menu to be
        ' removed completely.
        
        If Form1.List2.Selected(i) = True Then
            
             CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
             SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar", REG_DWORD, "1"

         Else
           
             Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar"
           
        End If

 
 
 Case 58:
 
        ' Remove common program groups from Start menu (All Versions)
        ' Disables the display of common groups when the
        ' user selects Programs from the Start menu.
        
        If Form1.List2.Selected(i) = True Then
            
             CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
             SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoCommonGroups", REG_DWORD, "1"

         Else
           
             Delete HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoCommonGroups"
           
        End If
 
 
           

   End Select
   
Next i


End Sub
