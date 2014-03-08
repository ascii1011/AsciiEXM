Attribute VB_Name = "mod_Drives"
Option Explicit

'0 = unknown
'1 = removable
'2 = fixed
'3 = network
'4 = cd
'5 = RAM disk

'oDrive.SerialNumber
'oDrive.ShareName
'oDrive.VolumeName
'oDrive.TotalSize
'oDrive.RootFolder
'oDrive.Path
'oDrive.IsReady
'oDrive.FileSystem
'oDrive.FreeSpace
'oDrive.DriveLetter
'oDrive.DriveType
Sub Drives2Struct()
    Dim oFS, oDrive As Drive
    Dim i As Integer
    Dim fr
    Dim ttl
    
On Error GoTo Err:

    Set oFS = CreateObject("scripting.filesystemobject")
    'Debug.Print oFS.Drives.Count
    ReDim Main.Drives(oFS.Drives.Count)
    
    frmExMerge.Output_2_List frmExMerge.List4, "Examining drives"
    
    For Each oDrive In oFS.Drives
    
        If oDrive.DriveLetter <> "" Then
        
            frmExMerge.Label50.Caption = "->" & oDrive.DriveLetter & ":\"
            Pause 0.02
            
            If oDrive.IsReady = True Then
            
                If oDrive.DriveType = 3 Then
                
                    'network
                    ttl = Drive_TotalSpace(Main.Drives(i).Letter, True)
                    fr = Drive_FreeSpace(Main.Drives(i).Letter, True)
                    
                    Main.Drives(i).Letter = oDrive.DriveLetter & ":\"
                    Main.Drives(i).Type = oDrive.DriveType 'Drive_Type(oDrive.DriveType)
                    Main.Drives(i).FileSystem = oDrive.FileSystem
                    Main.Drives(i).IsReady = oDrive.IsReady
                    Main.Drives(i).Path = oDrive.Path
                    Main.Drives(i).RootFolder = oDrive.RootFolder
                    Main.Drives(i).ShareName = oDrive.ShareName
                    Main.Drives(i).VolumeName = oDrive.VolumeName
                    Main.Drives(i).TotalSpace = ttl
                    Main.Drives(i).FreeSpace = fr
                    
                    i = i + 1
                ElseIf oDrive.DriveType = 4 Then
                
                Else
                
                    'all else - HardDisk, USB, CD
                    ttl = Drive_TotalSpace(Main.Drives(i).Letter, True)
                    fr = Drive_FreeSpace(Main.Drives(i).Letter, True)
                    
                    Main.Drives(i).Letter = oDrive.DriveLetter & ":\"
                    Main.Drives(i).Type = oDrive.DriveType 'Drive_Type(oDrive.DriveType)
                    Main.Drives(i).Serial = oDrive.SerialNumber
                    Main.Drives(i).TotalSpace = ttl
                    Main.Drives(i).FreeSpace = fr
                i = i + 1
                End If
            Else
                'not ready
                'Main.Drives(i).Letter = oDrive.DriveLetter & ":\"
                'Main.Drives(i).Type = oDrive.DriveType 'Drive_Type(oDrive.DriveType)
                'Main.Drives(i).IsReady = oDrive.IsReady
                'Main.Drives(i).Path = oDrive.Path
                'Main.Drives(i).ShareName = oDrive.ShareName
                
                '''''''''''Main.Drives(i).FileSystem = oDrive.FileSystem
                '''''''''''Main.Drives(i).RootFolder = oDrive.RootFolder
                '''''''''''Main.Drives(i).VolumeName = oDrive.VolumeName
                i = i + 1
            
            End If
        End If
        
    Next
    
    If i > 0 Then
        DisplayDrives
    End If
    
    Exit Sub
    
Err:
    MsgBox "Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description
    Resume Next
    'Exit Sub
End Sub




Sub DisplayDrives()
    Dim i As Integer
    
On Error GoTo Err:
    
    
    frmExMerge.MSFlexGrid2.Rows = UBound(Main.Drives) + 1 'add 1 for the column headers
    
    For i = 0 To UBound(Main.Drives) - 1
        If Main.Drives(i).Letter <> "" And Main.Drives(i).Type < 4 And Main.Drives(i).Type > 0 Then
            'lst.AddItem Main.Drives(i).Letter & _
                " [" & Main.Drives(i).FreeSpace & " free from " & _
                Main.Drives(i).TotalSpace & "] " & _
                Main.Drives(i).Type
                
            With frmExMerge.MSFlexGrid2
                .TextMatrix(i + 1, 1) = Main.Drives(i).Letter
                .TextMatrix(i + 1, 2) = Main.Drives(i).Path
                .TextMatrix(i + 1, 3) = Drive_Type(Main.Drives(i).Type)
                .TextMatrix(i + 1, 4) = Main.Drives(i).FreeSpace
                .TextMatrix(i + 1, 5) = Main.Drives(i).TotalSpace
            End With
        End If
    Next i
    
    Exit Sub
Err:
    Debug.Print "DisplayDrives-Err: " & Err.Number & ", Desc: " & Err.Description
    Resume Next
End Sub

