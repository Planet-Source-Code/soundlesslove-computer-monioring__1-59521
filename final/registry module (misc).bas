Attribute VB_Name = "Module8"


Public Sub GetRegistryValues2(n As Integer)


' this code on upon form activate gets the values for each
' of the options and set them...


    Dim i As Integer
    Dim value As Integer
    Dim val
    
       
    
For i = 0 To n - 1
     
  Select Case i
  
  
  Case 0:
                value = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "DisableScandiskOnBoot")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
        
        Case 1:
                value = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "Win31FileSystem")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 2:
                val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Sound", "Beep")
                If val = "No" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 3:
                value = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\CDRom", "Autorun")
                If value = 0 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 4:
                val = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\MediaPlayer\PlayerUpgrade", "AskMeAgain")
                If val = "no" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 5:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoRadioBar") And QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoMediaFavorites") And QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoFindNewStations")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 6:
                value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Symantec\Norton AntiVirus\Clinic", "DisableSplashScreen")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                
        Case 7:
                value = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Symantec\Norton Utilities", "DisableSplashScreen")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 8:
                
                
                
                val = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\Settings", "EnableDVDUI")
                If val = "Yes" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                
                
        Case 9:
                value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Telnet", "SmoothScroll")
                If value = 1 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
            
                
                
        Case 10:
                ' Use Windows Update Without Registering
                
                
                val = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Welcome\RegWiz", "RegDone")
                If val = "1" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                
                
                
        Case 11:
                val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "MinMaxClose")
                If val = "0" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                
                
                
        Case 12:
                val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "FontSmoothing")
                If val = "2" Then
                    Form1.List3.Selected(i) = True
                Else
                    
                    Form1.List3.Selected(i) = False
                 
                End If
                
                
        Case 13:
                val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size")
                If val = "16" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                
                
                
        Case 14:
                val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon BPP")
                If val = "16" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
                              
        Case 15:
         
               val = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "DragFullWindows")
               If val = "1" Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
                
        Case 16:
        
            value = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode")
            If value = 0 Then
                    Form1.List3.Selected(i) = True
                Else
                    Form1.List3.Selected(i) = False
                End If
            
         
         
        Case 17:
           val = QueryValue(HKEY_USERS, ".DEFAULT\Control Panel\Colors", "Background")
           If val = "0 0 0" Then
                    Form1.List3.Selected(i) = True
           Else
               Form1.List3.Selected(i) = False
           End If
           
           
        

Case 18:
         
        value = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NameNumericTail")
         If value = 1 Then
                    Form1.List3.Selected(i) = True
         Else
                   Form1.List3.Selected(i) = False
         End If
         
         
    Case 19
            value = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu")
            If value = 1 Then
                    Form1.List3.Selected(i) = True
            Else
                    Form1.List3.Selected(i) = False
            End If
                
                
                
    Case 20:
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\Open Find\Places\StandardPlaces", "Show")
            If value = 0 Then
                    Form1.List3.Selected(i) = True
            Else
                    Form1.List3.Selected(i) = False
                            
            End If

    Case 21:
            value = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\General", "AcbControl")
            If value = 1 Then
                    Form1.List3.Selected(i) = True
            Else
                    Form1.List3.Selected(i) = False
            End If
             
           End Select
    
Next i
  
 
End Sub





Public Sub Registry3(n As Integer)

' coding for the apply button on registry tab
' misc tab

Dim i As Integer
    
    
For i = 0 To n - 1
     
    Select Case i
      
      
        Case 0:
                ' Disable Scandisk After Bad Shut Down
                ' Normally if a system is shutdown
                ' improperly scandisk will run on the next
                ' reboot to ensure the contents of the hard
                ' disk are valid. This setting stops scandisk
                ' from running automatically when restarting.
                             
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem"
                SetKeyValue HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "DisableScandiskOnBoot", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "DisableScandiskOnBoot"
           
                End If
        
        
        Case 1:
                ' Control Long Filename Support (All Versions)
                ' Windows 9x and NT introduced the use of
                ' long filenames on existing FAT partitions.
                ' Some legacy software maybe incompatible
                ' with this new file system design, and may
                ' require the use of 8.3 filenames.
                ' By enabling this setting you can turn off
                ' long filename support.
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem"
                SetKeyValue HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "Win31FileSystem", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\FileSystem", "Win31FileSystem"
           
                End If
                
        Case 2:
                ' Disable PC Speaker Beeping on Errors
                ' If you get annoyed by the beeps and
                ' noises coming from your PC speaker but
                ' can't find a way to turn it off, then use
                ' this tip to disable it.
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Control Panel\Sound"
                SetKeyValue HKEY_CURRENT_USER, "Control Panel\Sound", "Beep", REG_SZ, "No"

                Else
           
                Delete HKEY_CURRENT_USER, "Control Panel\Sound", "Beep"
           
                End If
                
        Case 3:
                ' Disable Auto Run For CD-ROMS
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\CDRom"
                SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\CDRom", "Autorun", REG_DWORD, "0"

                Else
           
                SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\CDRom", "Autorun", REG_DWORD, "1"
           
                End If
                
                
        Case 4:
                ' Disable Media Player Upgrade Message
                ' This setting allows you to disable the
                ' automatic Media Player upgrade message that
                ' appears when a new version of the player has
                ' been released.
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\MediaPlayer\PlayerUpgrade"
                SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\MediaPlayer\PlayerUpgrade", "AskMeAgain", REG_SZ, "no"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\MediaPlayer\PlayerUpgrade", "AskMeAgain"
           
                End If
                
                
        Case 5:
                ' Remove Items from Media Player (All Versions)
                ' This tweak allows you to remove the Radio
                ' Bar, Media Favorites and Finding New
                ' Station from Windows Media Player.
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoRadioBar", REG_DWORD, "1"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoMediaFavorites", REG_DWORD, "1"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoFindNewStations", REG_DWORD, "1"

                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoRadioBar"
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoMediaFavorites"
                Delete HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\WindowsMediaPlayer", "NoFindNewStations"
           
                End If
                
                
        Case 6:
                ' Disable the Norton AntiVirus Splash Screen (All Versions)
                ' This setting will stop the Norton
                ' AntiVirus Scanner startup logo from being
                ' shown.
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Symantec\Norton AntiVirus\Clinic"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Symantec\Norton AntiVirus\Clinic", "DisableSplashScreen", REG_DWORD, "1"
                
                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Symantec\Norton AntiVirus\Clinic", "DisableSplashScreen"
                
                End If
                
                
                
        Case 7:
                ' Disable the Norton Utilities Splash Screen
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Symantec\Norton Utilities"
                SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Symantec\Norton Utilities", "DisableSplashScreen", REG_DWORD, "1"
                
                Else
           
                Delete HKEY_LOCAL_MACHINE, "SOFTWARE\Symantec\Norton Utilities", "DisableSplashScreen"
                
                End If
                
                
        Case 8:
                ' Enable DVD Features in Media Player
                
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\Settings"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\Settings", "EnableDVDUI", REG_SZ, "Yes"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\Player\Settings", "EnableDVDUI"
                
                End If
                
                
        Case 9:
                ' Use Smooth Scrolling in Telnet
                
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Telnet"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Telnet", "SmoothScroll", REG_DWORD, "1"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Telnet", "SmoothScroll"
                
                End If
                
                
        Case 10:
                ' Use Windows Update Without Registering
                
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Welcome\RegWiz"
                SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Welcome\RegWiz", "RegDone", REG_SZ, "1"
                
                Else
           
                Delete HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Welcome\RegWiz", "RegDone"
                
                End If
                
                
        Case 11:
                ' Remove the Minimize, Maximize and Close
                ' Tooltips
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Control Panel\Desktop"
                SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "MinMaxClose", REG_SZ, "0"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Control Panel\Desktop", "MinMaxClose"
                
                End If
                
                
        Case 12:
                ' Enable Font Smoothing
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Control Panel\Desktop"
                SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "FontSmoothing", REG_SZ, "2"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Control Panel\Desktop", "FontSmoothing"
                
                End If
                
        Case 13:
                ' Change the Size of Desktop Icons
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics"
                SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", REG_SZ, "16"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size"
                
                End If
                
                
        Case 14:
                ' Displaying Hi-Color Icons without the
                ' Plus Pack
                
                If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics"
                SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", REG_SZ, "16"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon BPP"
                
                End If
                
        Case 15:
         
         'This setting enables the Full Windows Drag function
         ', which allow you to view the contents of window while dragging it across the screen, instead of
         'just the standard outline.
        
       
         If Form1.List3.Selected(i) = True Then
           
            CreateNewKey HKEY_CURRENT_USER, "Control Panel\Desktop"
           SetKeyValue HKEY_CURRENT_USER, "Control Panel\Desktop", "DragFullWindows", REG_SZ, "1"
           

         Else
         
           Delete HKEY_CURRENT_USER, "Control Panel\Desktop", "DragFullWindows"
           
         End If
        Case 16:
        
           'Normally when the contents of a window change you may need to wait a few seconds, or press F5, to refresh the display to see the updated information. This tweak configures the system to perform faster automatic updates.
         If Form1.List3.Selected(i) = True Then
           
            CreateNewKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Update"
            SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode", REG_DWORD, 0
         Else
         
           SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode", REG_DWORD, 1
         End If
        Case 17:
         'When you change the color scheme and appearance of your desktop
'it does not change the background color of the logon screen to match. This tweak allows you to change that color as well.

            If Form1.List3.Selected(i) = True Then
           
            CreateNewKey HKEY_USERS, ".DEFAULT\Control Panel\Colors"
            SetKeyValue HKEY_USERS, ".DEFAULT\Control Panel\Colors", "Background", REG_SZ, "0 0 0"

         Else
         
           Delete HKEY_USERS, ".DEFAULT\Control Panel\Colors", "Background"
           
         End If
   
Case 18:
  'When long filenames are shown in an application that only supports short filenames a tilde "~" is used to convert the long name into a compatible short name. This setting removes the use of tildes
   
   If Form1.List3.Selected(i) = True Then
           
            CreateNewKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem"
            SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NameNumericTail", REG_DWORD, 1

         Else
         
           SetKeyValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NameNumericTail", REG_DWORD, 0
           
         End If
         
    Case 19:
            ' Remove the Favorites Folder from the Start Menu
            
            If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", REG_DWORD, "1"
                
                Else
           
                Delete HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu"
                
                End If
                
    Case 20:
            ' Remove Items from Office Places Bar
            
            If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\Open Find\Places\StandardPlaces"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\Open Find\Places\StandardPlaces", "Show", REG_DWORD, "0"
                
                Else
           
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\Open Find\Places\StandardPlaces", "Show", REG_DWORD, "1"
                
                End If

    Case 21:
            ' Prevent the Office Clipboard Toolbar
            'from Appearing
            
            If Form1.List3.Selected(i) = True Then
            
                CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\General"
                SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\General", "AcbControl", REG_DWORD, "1"
                
                Else
           
                Delete HKEY_CURRENT_USER, "Software\Microsoft\Office\9.0\Common\General", "AcbControl"
                
                End If
 
    End Select
    
Next i

End Sub




