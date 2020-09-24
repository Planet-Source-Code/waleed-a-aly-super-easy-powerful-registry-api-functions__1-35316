Attribute VB_Name = "modRegistryAPI"
'*******************************************************'
'                                                       '
'   By:         Waleed A. Aly                           '
'   ASL:        [20 M Egypt]                            '
'   eMail:      wa_aly@tdcspace.dk                      '
'   Thanks to:  www.allapi.net                          '
'                                                       '
'     Please eMail me any Comments and|or Suggestions.  '
'   I hope you like my work and think is usefull !  :)  '
'   but please Notify me with any [Enhancements] you    '
'   apply to this code. Also i'd love to know how many  '
'   people are using my Code so you can always eMail me '
'   if you are goin' to use it :)                       '
'                                      Thanks.          '
'                                                       '
'*******************************************************'

Option Explicit

'Status (response) for the [CreateRegKey] function
Public Enum Status
    Key_Created = 0                   ' Key successfully created
    Key_Already_Exists = 1            ' Key already exists
    Key_Could_NOT_be_Created = 2      ' Error creating key (can't create)
End Enum

'Handle for the manipulated key
Private Handle As Long

'StartUp key Entry located in the [HKEY_LOCAL_MACHINE]
Public Const StartUpEntry As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

'returned by API (Disposition)
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'predefined Root Keys (Reserved Handles)
Public Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

'Value Types
Public Enum ValueType
    REG_SZ = 1                        ' fixed-length text string
    REG_EXPAND_SZ = 2                 ' variable-length data string (used for variables that are resolved when a program or service uses the data)
    REG_BINARY = 3                    ' Binary Data
    REG_DWORD = 4                     ' 4-Bytes long Data
    REG_MULTI_SZ = 7                  ' multiple string (Lists or Multiple Values: separated by Spaces, Commas, or other Marks)
    REG_FULL_RESOURCE_DESCRIPTOR = 9  ' series of nested arrays designed to store a resource list for a hardware component or driver
End Enum

'Key Options
Public Enum Options
    REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
    REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
    REG_OPTION_BACKUP_RESTORE = 4     ' Opens for backup or restore (overrides security options)
End Enum

'Other Constants (Needed)
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

'Key Access Security Options
Public Enum AccessMethod
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
End Enum

'API Registry Functions
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Function CreateRegKey(ByVal hKey As hKey, ByVal SubKey As String, ByVal KeyOptions As Options, ByVal AccessPermission As AccessMethod) As Status
'parameter [SubKey] must not begin with the backslash character "\" and cannot be [NULL]
'Call the function [GetKeyHandle As Long] to get the Handle for the Created Key
'function returns
' 0 : if successfully created          (the Key Handle is retrieved)
' 1 : if key already exists            (the Key Handle is retrieved)
' 2 : if key could not be created      (Previous Key Handle is LOST)

    Dim Disposition As Long
    
    Call RegCreateKeyEx(hKey, SubKey, 0, "", KeyOptions, AccessPermission, ByVal 0, Handle, Disposition)
    
    If Disposition = REG_CREATED_NEW_KEY Then
        CreateRegKey = Key_Created
    ElseIf Disposition = REG_OPENED_EXISTING_KEY Then
        CreateRegKey = Key_Already_Exists
    ElseIf Handle = 0 Then
        CreateRegKey = Key_Could_NOT_be_Created
    End If

End Function

Public Function OpenRegKey(ByVal hKey As hKey, ByVal SubKey As String, ByVal AccessPermission As AccessMethod) As Boolean
'returns [TRUE] at success
'returns [FALSE] if Key Does [not EXIST] or cannot be Opened (Previous Key Handle is LOST)
'function returns the key handle for the opened key as the variable [Handle]
'Call the function [GetKeyHandle As Long] to get the Handle for the Opened Key

    If RegOpenKeyEx(hKey, SubKey, 0, AccessPermission, Handle) = 0 Then OpenRegKey = True

End Function

Public Function CloseRegKey(ByVal KeyHandle As Long) As Boolean
'returns [TRUE] at success
'returns [FALSE] if Key could not be closed (KeyHandle is NOT a Registry Key Handle)
'the closed Key Handle is Reset. [Handle = 0]

    If RegCloseKey(KeyHandle) = 0 Then CloseRegKey = True
    Handle = 0

End Function

Public Function DeleteRegKey(ByVal hKey As hKey, ByVal SubKey As String) As Boolean
'[SubKey] parameter cannot be [NULL]
'specified Key to delete cannot contain [SUBKEYS]
'returns [TRUE] at success
'returns [FALSE] if Key could not be deleted

    If RegDeleteKey(hKey, SubKey) = 0 Then DeleteRegKey = True

End Function

Public Function SetRegValue(ByVal KeyHandle As Long, ByVal ValueName As String, ByVal ValueType As ValueType, ByVal ValueData As Variant) As Boolean
'returns [TRUE] at success
'returns [FALSE] if Value could not be created
'if Value does not already exist in the specified Key, it is added
'this function supports setting [String] & Long (4Bytes) [Binary] Values
'for floating-point Numbers or Integers exceeding the range of Long, save them as string Values

    Select Case ValueType
        Case REG_SZ
            If RegSetValueEx(KeyHandle, ValueName, 0, ValueType, ByVal CStr(ValueData), Len(ValueData)) = 0 Then SetRegValue = True
        Case REG_BINARY
            If RegSetValueEx(KeyHandle, ValueName, 0, ValueType, CLng(Val(ValueData)), 4) = 0 Then SetRegValue = True
    End Select

End Function

Public Function QueryRegValue(ByVal KeyHandle As Long, ByVal ValueName As String) As String
'this function supports retrieving [String] & Long (4Bytes) [Binary] Values
'returns a [String] containing whatever Value is retrieved

    Dim ValueType As ValueType, BufferSize As Long
    Dim StringBuffer As String, BinaryBuffer As Long
    
    If RegQueryValueEx(KeyHandle, ValueName, 0, ValueType, ByVal 0, BufferSize) = 0 Then
        Select Case ValueType
            Case REG_SZ
                StringBuffer = String(BufferSize, Chr(0))
                Call RegQueryValueEx(KeyHandle, ValueName, 0, 0, ByVal StringBuffer, BufferSize)
                QueryRegValue = Left(StringBuffer, InStr(1, StringBuffer, Chr(0)) - 1)
            Case REG_BINARY
                Call RegQueryValueEx(KeyHandle, ValueName, 0, 0, BinaryBuffer, BufferSize)
                QueryRegValue = Str(BinaryBuffer)
        End Select
    End If

End Function

Public Function DeleteRegValue(ByVal KeyHandle As Long, ByVal ValueName As String) As Boolean
'returns [TRUE] at success
'returns [FALSE] if Value could not be Deleted or does Not Exist

    If RegDeleteValue(KeyHandle, ValueName) = 0 Then DeleteRegValue = True

End Function

Public Function GetKeyHandle() As Long
'returns [CURRENT] Key Handle [As Long] (i.e. handle for last Opened or Created Key)
'you should save KeyHandles in your program as you open them, if you'll be using many at the same time

    GetKeyHandle = Handle

End Function
