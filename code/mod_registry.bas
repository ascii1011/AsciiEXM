Attribute VB_Name = "mod_registry"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const HKEY_CURRENT_CONFIG = &H80000004

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Const KEY_READ = &H20019
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234

Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)


Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
    ' Description:
    ' This Function will Delete a key
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    ' Description:
    ' This Function will delete a value
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteValue(hKey, sValueName)
    RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
        
    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    
    On Error GoTo QueryValueExError
    
    'find cch & lType (string, Dword, etc)
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
    
    Select Case lType
        Case REG_SZ:
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch)
        Else
            vValue = Empty
        End If
    Case REG_DWORD:
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
    Case Else
        lrc = -1
    End Select
    
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
    
QueryValueExError:
    Resume QueryValueExExit

End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    ' Description:
    ' This Function will create a new key
    Dim hNewKey As Long
    Dim lRetVal As Long
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    ' Description:
    ' This Function will set the data field of a value
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    ' Description:
    ' This Function will return the data field of a value
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function
Function CheckRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
    Dim handle As Long
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_ALL_ACCESS, handle) = 0 Then
        CheckRegistryKey = True
        RegCloseKey handle
    End If
End Function


Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal As Long
    
    Set EnumRegistryValues = New Collection
    
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        hKey = handle
    End If
    
    Do
        nameLen = 260
        name = Space$(nameLen)
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        If retVal Then Exit Do
        
        valueInfo(0) = Left$(name, nameLen)
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
        End Select
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        index = index + 1
    Loop
    If handle Then RegCloseKey handle
        
End Function




Function ListRegistryKeyValuesInALocation(ByVal hKey As Long, ByVal KeyName As String) As String
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim valueInfoAll(5) As Variant
    Dim retVal As Long
    Dim sResult As String
    
    'Set EnumRegistryValues = New Collection
    sResult = ""
    
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        hKey = handle
    End If
    
    Do
        nameLen = 260
        name = Space$(nameLen)
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        If retVal Then Exit Do
        
        valueInfo(0) = Left$(name, nameLen)
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
        End Select
        'EnumRegistryValues.Add valueInfo, valueInfo(0)
        If valueInfo(0) <> "" And valueInfo(1) <> "" Then
            sResult = sResult & "[" & valueInfo(0) & "]" & " = " & valueInfo(1) & vbNewLine
        End If
        index = index + 1
    Loop
    ListRegistryKeyValuesInALocation = sResult
    If handle Then RegCloseKey handle
        'MsgBox valueInfoAll(0) & vbNewLine & valueInfoAll(1)
End Function










