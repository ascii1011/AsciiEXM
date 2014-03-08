Attribute VB_Name = "mod_system_info"
Option Explicit

Public sUser        As String   'using
Public sComputer    As String   'using
Public sOS          As String   'using
Public sBuild     As String
Public sVMinor      As String
Public sVMajor      As String
Public sVersion     As String
Public sGRootDir    As String
Public sMemory      As String

Public iOS          As String   'using
Public sIP          As String   'using
Public sMac         As String
Public sCPU         As String
Public strIP        As String
Public strName      As String


            
'''''''''''''''''''''''get operating system info''''''''''''''''''''''''''

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'''''''''''mac address'''''''''''''''''''''''''''
Public Const NO_ERROR = 0

Declare Function Inet_addr Lib "wsock32.dll" _
  Alias "inet_addr" (ByVal s As String) As Long

Declare Function SendARP Lib "iphlpapi.dll" _
  (ByVal DestIP As Long, _
   ByVal SrcIP As Long, _
   pMacAddr As Long, _
   PhyAddrLen As Long) As Long

Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   src As Any, _
   ByVal bcount As Long)
'''''''''''''''''''''''''''''''Grab current user name'''''''''''''''''''''''''''''''''
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
            (ByVal lpBuffer As String, nSize As Long) As Long
            '''''''''''''''''''''''get operating system info''''''''''''''''''''''''''
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
             (LpVersionInformation As OSVERSIONINFO) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



'''''''''''''''''''''''get operating system info''''''''''''''''''''''''''
Sub Get_OS_Info()
    
    Dim NewLine As String     ' New-line.
    Dim ret As Integer        ' OS Information
    Dim ver_major As Integer  ' OS Version
    Dim ver_minor As Integer  ' Minor Os Version
    Dim Build As Long         ' OS Build
    
On Error Resume Next
    
    NewLine = vbNewLine 'Chr(13) + Chr(10)  ' New-line.

    iOS = 0
    
    ' Get operating system and version.
    Dim verinfo As OSVERSIONINFO
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)
    If ret = 0 Then
        MsgBox "Error Getting Version Information"
        End
    End If
    
    ver_major = verinfo.dwMajorVersion
    ver_minor = verinfo.dwMinorVersion
    Build = verinfo.dwBuildNumber
       
    sVMinor = ver_minor
    sVMajor = ver_major
    sBuild = Build
    sVersion = Build & ", ver. " & ver_major & "." & ver_minor
    If verinfo.dwPlatformId = 2 And verinfo.dwMajorVersion = 5 And verinfo.dwMinorVersion = 1 And verinfo.dwBuildNumber = 2600 Then
        sOS = "XP"
        iOS = 1
        sGRootDir = "WINDOWS"
    Else
        If verinfo.dwPlatformId = 2 And verinfo.dwMajorVersion = 5 And verinfo.dwMinorVersion = 0 And verinfo.dwBuildNumber = 2195 Then
            sOS = "2K"
            iOS = 2
            sGRootDir = "WINNT"
        Else
            sOS = "Other"
            iOS = 10
        End If
    End If
End Sub

Sub SysInfo()
    Get_OS_Info
    Get_User_Name
    Get_Computer_name
End Sub

'''''''''''''''''''''''''''''''Grab current user name'''''''''''''''''''''''''''''''''

            
Sub Get_User_Name()
    ' Dimension variables
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

    ' Display the User Name
    sUser = UserName
    
End Sub


Sub Get_Computer_name()
    Dim strBuffer As String
    Dim lngBufSize As Long
    Dim lngStatus As Long
  
    lngBufSize = 255
    strBuffer = String$(lngBufSize, " ")
    lngStatus = GetComputerName(strBuffer, lngBufSize)
    If lngStatus <> 0 Then
        Main.Sys.ComputerName = LCase(Trim(Left(strBuffer, lngBufSize)))
    End If
End Sub





