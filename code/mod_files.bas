Attribute VB_Name = "mod_files"
Option Explicit


'Private fso As New Scripting.FileSystemObject

Function Exists_fso(sFilePath As String) As Boolean
    Dim fso
    Set fso = New Scripting.FileSystemObject
    Exists_fso = False
On Error Resume Next

    If fso.FileExists(sFilePath) Then Exists_fso = True
End Function

Public Function Exists_dir(FileName As String) As Boolean
    Exists_dir = Dir(FileName) <> ""
End Function

Function DelFile(sTarget As String) As Boolean
    Dim vFile
    Dim fso
    Set fso = New Scripting.FileSystemObject
    
    DelFile = False
On Error GoTo Err:

    Set vFile = fso.GetFile(sTarget)
    vFile.Delete
    DelFile = True

    Exit Function
Err:
    Debug.Print Err.Number & " " & Err.Description
End Function

Function Delete_File(sTarget As String) As Boolean
    Dim fso As FileSystemObject
    
    Delete_File = False
On Error GoTo Err:
    'sTarget = LCase(sTarget)
    'If fso.FileExists(sTarget) = True Then
        fso.DeleteFile sTarget, True
    'End If

    Exit Function
Err:
    Debug.Print Err.Number & " " & Err.Description
    Exit Function
End Function

Function Move_File(sFrom As String, sTo As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Move_File = False
On Error GoTo Err:

    'If fso.FileExists(sFrom) = True Then
        'If fso.FolderExists(sTo) = True Then
            fso.MoveFile sFrom, sTo
            Move_File = True
        'End If
    'End If
    Exit Function
Err:
    Debug.Print Err.Number & " " & Err.Description
    Exit Function
End Function

Function Rename_File(sFrom As String, sTo As String) As Boolean
    Dim fso As FileSystemObject, fso_File
    Set fso = CreateObject("scripting.filesystemobject")
    
    Debug.Print sFrom & " - " & sTo
    
    Rename_File = False
On Error GoTo Err:

    Name sFrom As sTo
    'Set fso_File = fso.GetFile(sFrom)
    'fso_File.name sFrom, sTo, True
    
    Rename_File = True
    Exit Function
Err:
    Debug.Print Err.Number & " " & Err.Description
    Exit Function
End Function

Function Copy_File(sFrom As String, sTo As String) As Boolean
    Dim fso As FileSystemObject
    Set fso = CreateObject("scripting.filesystemobject")
    
    Debug.Print sFrom & " - " & sTo
    
    Copy_File = False
On Error GoTo Err:

    'sTo = sTo & "\MailBoxes.zip"
    'If fso.FileExists(LCase(sFrom)) = True Then
        'If fso.FolderExists(sTo) = True Then
            fso.CopyFile sFrom, sTo, True
        'End If
    'End If
    Copy_File = True
    Exit Function
Err:
    Debug.Print Err.Number & " " & Err.Description
    Exit Function
End Function

