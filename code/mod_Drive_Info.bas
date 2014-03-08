Attribute VB_Name = "mod_Drive_Info"

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This sample uses one form with one command button and
' three text boxes.
' Type the drive letter(or UNC) you want to know the size
' of into Text1.
' Text2 returns the total number of bytes on the drive.
' Text3 returns the total number of free bytes on the drive.
' When using an API call that expects an unsigned long
' integer you need to pass it a currency datatype to capture
' the value because VB does not have an unsigned long.
' However the currency will have the value and can be
' converted once returned.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''Drive space information
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
          Alias "GetDiskFreeSpaceExA" _
          (ByVal lpDirectoryName As String, _
          lpFreeBytesAvailableToCaller As Currency, _
          lpTotalNumberOfBytes As Currency, _
          lpTotalNumberOfFreeBytes As Currency) As Long
          
Private Const lGigaByte As Long = 1073741824
Private Const lMegaByte As Long = 1048576

''''''''''''''''Drive Type
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_REMOVABLE = 1
Private Const DRIVE_FIXED = 2
Private Const DRIVE_REMOTE = 3
Private Const DRIVE_CDROM = 4
Private Const DRIVE_RAMDISK = 5

'0 = unknown
'1 = removable
'2 = fixed
'3 = network
'4 = cd
'5 = RAM disk

Function Drive_Type(sDriveLetter As String) As String
    Select Case sDriveLetter
        Case DRIVE_UNKNOWN
            Drive_Type = "Unknown"
        Case DRIVE_REMOVABLE
            Drive_Type = "removable drive"
        Case DRIVE_FIXED
            Drive_Type = "fixed drive"
        Case DRIVE_REMOTE
            Drive_Type = "remote drive"
        Case DRIVE_CDROM
            Drive_Type = "CD-ROM drive"
        Case DRIVE_RAMDISK
            Drive_Type = "RAM disk"
    End Select
End Function

Public Function Drive_TotalSpace(sDriveLetter As String, ByVal bFormat As Boolean) As String
     Dim cTmp As Currency
     Dim cTotal As Currency     'cTotalNumberOfBytesOnDrive
     Dim cFree As Currency      'cTotalNumberOfFreeBytes
     Dim lResults As Long

     lResults = GetDiskFreeSpaceEx(sDriveLetter, cTmp, cTotal, cFree)
                    
     Drive_TotalSpace = Format _
          ((cTotal * 10000) / IIf(cTotal * 10000 >= lGigaByte, _
                lGigaByte, lMegaByte), IIf(cTotal * 10000 >= lGigaByte, "##0.## GB", "##0.## MB"))
         
End Function

Public Function Drive_FreeSpace(sDriveLetter As String, ByVal bFormat As Boolean) As String
     Dim cTmp As Currency
     Dim cTotal As Currency     'cTotalNumberOfBytesOnDrive
     Dim cFree As Currency      'cTotalNumberOfFreeBytes
     Dim lResults As Long

     lResults = GetDiskFreeSpaceEx(sDriveLetter, cTmp, cTotal, cFree)
                    
     Drive_FreeSpace = Format _
          ((cFree * 10000) / IIf(cFree * 10000 >= lGigaByte, lGigaByte, lMegaByte), _
                IIf(cFree * 10000 >= lGigaByte, "##0.## GB", "##0.## MB"))
End Function


'Public Sub DriveInfo()
'    ' Dimension local variables.
'     Dim cJunk As Currency
'     Dim cTotalNumberOfBytesOnDrive As Currency
'     Dim cTotalNumberOfFreeBytes As Currency
'     Dim lResults As Long

'     lResults = GetDiskFreeSpaceEx("c:\", cJunk, _
'                    cTotalNumberOfBytesOnDrive, cTotalNumberOfFreeBytes)
     ' Format and Display the TotalNumberOfBytesOnDrive value in
     ' GB or MB depending on size.
     ' Multiply the TotalNumberOfBytesOnDrive value by 10000 to
     ' convert the currency data into a long.
'     Form6.List1.AddItem Format _
'          ((cTotalNumberOfBytesOnDrive * 10000) / _
'          IIf(cTotalNumberOfBytesOnDrive * 10000 >= lGigaByte, _
'          lGigaByte, lMegaByte), IIf(cTotalNumberOfBytesOnDrive * _
'         10000 >= lGigaByte, "##0.###GB", "##0.###MB"))
     ' Format and Display the TotalNumberOfFreeBytes value in
     ' GB or MB depending on size. Remember 1 Gigabyte 1024 Megabytes.
     ' Multiply the TotalNumberOfFreeBytesvalue by 10000 to convert
     ' the currency data into a long.
'     Form6.List1.AddItem Format _
'          ((cTotalNumberOfFreeBytes * 10000) / IIf(cTotalNumberOfFreeBytes * _
'            10000 >= lGigaByte, lGigaByte, lMegaByte), _
'             IIf(cTotalNumberOfFreeBytes * 10000 >= lGigaByte, "##0.###GB", _
'            "##0.###MB"))
'End Sub





