Attribute VB_Name = "mod_functinos"
Option Explicit

'shell and wait variables'''''''''''''''''
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const Infinite = -1&
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
'Private Const MAX_PATH As Long = 260


Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle) As Boolean
    Dim process_id As Long
    Dim process_handle As Long
    
    ShellAndWait = False
    
On Error GoTo ShellError
    
    DoEvents
    If bHaltProcess = True Then Exit Function
    
    process_id = Shell(program_name, window_style)
    
On Error GoTo 0
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, Infinite
        CloseHandle process_handle
    End If
    
    ShellAndWait = True
    
    Exit Function

ShellError:
    Exit Function
End Function

Function funCreatefile(sFilePath As String, sBody As String) As Boolean
    Dim fs, a

    funCreatefile = False
On Error GoTo errHandler
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(sFilePath, True)
    a.Write (sBody)
    a = FreeFile
    funCreatefile = True
    Exit Function
    
errHandler:
    a = FreeFile
End Function


Sub Pause(Seconds)
    Dim PauseTime, Start1, finish1
    PauseTime = Seconds   ' Set duration.
    Start1 = Timer   ' Set start time.
    Do While Timer < Start1 + Seconds
        DoEvents    ' Yield to other processes.
    Loop
    finish1 = Timer  ' Set end time.
End Sub


Public Function FileExists(FileName As String) As Boolean
    FileExists = Dir(FileName) <> ""
End Function

Function FileCopy(sSource As String, sDestination As String) As Boolean
    
End Function



Function comDialog_Folder() As Object
    Dim shl As Shell, hw As Long, sMsg As String, opt As Long
    Set shl = New Shell
    
    hw = frmExMerge.hWnd
    sMsg = "Select a Folder"
    opt = BIF_RETURNONLYFSDIRS
    'MsgBox shl.Namespace.GetDetailsOf
    Set comDialog_Folder = shl.BrowseForFolder(hw, sMsg, opt)
    
    Dim i
    For i = 0 To comDialog_Folder.Items.Count
        Debug.Print comDialog_Folder.Items.Item(i)
    Next i
    
    
    MsgBox comDialog_Folder.ParseName
    
End Function

Function SHFolder() As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = frmExMerge.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    SHFolder = sPath
End Function

'Filter = "All Files .* |*.*"
Function comDialog_File(sFilter) As String
    Dim cdObj As CommonDialog
    
    cdObj.Filter = sFilter
    'CommonDialog1.InitDir = Trim(frmExMerge.txtPSTPath.Text)
    cdObj.ShowOpen
    
    If cdObj.FileName = "" Then
        Exit Function
    End If
    
    comDialog_File = cdObj.FileName
    
End Function
