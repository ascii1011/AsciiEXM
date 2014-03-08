Attribute VB_Name = "mod_winzip32"
Option Explicit

Function ZipAFile(sDestinationPath As String, sSourcePath As String) As Boolean
    Dim WinZip32_Executable As String
    Dim String_ToBe_Executed As String
    
    ZipAFile = False

    WinZip32_Executable = "C:\Program Files\WinZip\winzip32"
    
    If FileExists(WinZip32_Executable & ".exe") Then
        String_ToBe_Executed = WinZip32_Executable & " -a " & sDestinationPath & " " & sSourcePath
    End If
        
    ZipAFile = ShellAndWait(String_ToBe_Executed, vbNormal)
        
End Function


