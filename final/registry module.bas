Attribute VB_Name = "Module3"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' constatnts declared for api functions...

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_NONE = 0                       ' No value type
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Public Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Public Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' api functions declared for various tasks...

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
 "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
 String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
 As String, lpcbData As Long) As Long
   
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long
   
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long
   
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long
   
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
ByVal cbData As Long) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' original api funtion... for setting a value in registry..

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   Public Function SetValueEx(ByVal hKey As Long, svaluename As String, _
   lType As Long, vValue As Variant) As Long
   
       Dim lValue As Long
       Dim sValue As String
       
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hKey, svaluename, 0&, _
                                              lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValueEx = RegSetValueExLong(hKey, svaluename, 0&, _
   lType, lValue, 4)
   
           End Select
           
   End Function

   
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   
   ' original api function for querying the value from registry...
   
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   
   Public Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
       
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
       
   End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  function made by me to make easy use of original api funtion...
' usage give first predefined key
' then the path "include folders to create also"
' this funtion only creates the folder in hiearchy of registry...

''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub CreateNewKey(lpredefinedkey As Long, sNewKeyName As String)
       
       Dim hNewKey As Long         'handle to the new key
       Dim lRetVal As Long         'result of the RegCreateKeyEx function

       lRetVal = RegCreateKeyEx(lpredefinedkey, sNewKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
                 
       RegCloseKey (hNewKey)
       
       
   End Sub
   
   
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  function made by myself ...
' it creates the value in the newly created folder or
' in the previos folder...
' usage 1. give predefined section in capital letters...
' 2. give the exact path of the key (include all folders)
' 3. give the name of the key to be craeted (any)
' 4. give its type. (string or dword...or anything.) .
'    remembre to include its constant from apiviewer ...
' 5. finally give its value. in case of dword its value must be in decimals...


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub SetKeyValue(lpredefinedkey As Long, skeyname As String, svaluename As String, _
   lValueType As Long, vValueSetting As Variant)
       
       
       Dim lRetVal As Long      'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key
       
       lRetVal = RegOpenKeyEx(lpredefinedkey, skeyname, 0, _
                                 KEY_SET_VALUE, hKey)
                                 
       lRetVal = SetValueEx(hKey, svaluename, lValueType, vValueSetting)
       
       RegCloseKey (hKey)
       
       
   End Sub
   
   


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  function made by myself ...
' it returns the value of the newly created key or existing key
' in the folder...
' usage 1. give predefined section in capital letters...
' 2. give the exact path of the key (include all folders)
' 3. give the name of the key for which u want its value (must be there). else error code (badkey)...
' 4. it returns the value of type variant from where it is called...


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Public Function QueryValue(lpredefinedkey As Long, skeyname As String, svaluename As String)

       
       Dim lRetVal As Long      'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant    'setting of queried value

       lRetVal = RegOpenKeyEx(lpredefinedkey, skeyname, 0, _
   KEY_QUERY_VALUE, hKey)
   
    '  MsgBox lRetVal
      
       lRetVal = QueryValueEx(hKey, svaluename, vValue)
       
       'MsgBox vValue
       QueryValue = vValue
              
       RegCloseKey (hKey)
       
   End Function


' function which deletes the specific value from the registry...

Public Sub Delete(lpredefinedkey As Long, skeyname As String, svaluename As String)

Dim lRetVal As Long      'result of the regopenkey function
Dim hKey As Long         'handle of open key

lRetVal = RegOpenKeyEx(lpredefinedkey, skeyname, 0, KEY_SET_VALUE, hKey)

lRetVal = RegDeleteValue(hKey, svaluename)

RegCloseKey hKey

End Sub

' funtion which deletes a complete folder in the registry...

Public Sub DeleteFolder(lpredefinedkey As Long, skeyname As String, svaluename As String)

Dim lRetVal As Long      'result of the regopenkey function
Dim hKey As Long         'handle of open key

lRetVal = RegOpenKeyEx(lpredefinedkey, skeyname, 0, KEY_SET_VALUE, hKey)

lRetVal = RegDeleteKey(hKey, svaluename)

RegCloseKey hKey

End Sub

