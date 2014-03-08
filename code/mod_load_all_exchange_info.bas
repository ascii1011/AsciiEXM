Attribute VB_Name = "mod_load_all_exchange_info"
Option Explicit

Private oExchSVR As New CDOEXM.ExchangeServer
Private arrSGs() As CDOEXM.StorageGroup
Private arrEXSG As Variant

'Dim oExchangeDB As New CDOexm.MailBoxStoreDB
'Dim arrMSs() As CDOexm.MailBoxStoreDB

'Dim oMBX As CDOexm.IExchangeMailbox

Public iServers As Integer
Private iStorageGroups As Integer
Private iMailboxStoreDB As Integer
Private iMailBoxes As Integer


Sub GetAllExchangeInfo()

    GetExchangeServers_Main
    frmExMerge.Label10.Caption = ""
    frmExMerge.Label50.Caption = ""
    frmExMerge.Label60.Caption = ""
End Sub

Function GetExchangeServers_Main()
    Dim iAdRootDSE As IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim varConfigNC As Variant
    Dim strQuery As String
        
On Error GoTo errHandler:
        
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    varConfigNC = iAdRootDSE.Get("configurationNamingContext")

    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"

    strQuery = "<LDAP://" & varConfigNC & ">;(objectCategory=msExchExchangeServer);name,serialNumber;subtree"

    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    If Not Rs.EOF Then
        iServers = 0
        
        Rs.MoveFirst
        
        Main.Exch.ServerCount = Rs.RecordCount
        ReDim Main.Exch.Svrs(Main.Exch.ServerCount)
    
        While Not Rs.EOF
                        
            Main.Exch.Svrs(iServers).name = Rs.Fields("name")
            
    On Error GoTo SvrSkip:
            
            frmExMerge.List4.AddItem "Exchange '" & Rs.Fields("name") & "' discovery ..."
            frmExMerge.Label10.Caption = "Exchange '" & Rs.Fields("name") & "' discovery ..."
            frmExMerge.Output_2_List frmExMerge.List4, "Digging Exchange Server '" & Rs.Fields("name") & "'"
            
            Pause 1
            oExchSVR.DataSource.Open Rs.Fields("name")
            
            Main.Exch.Svrs(iServers).DaysBeforeLogFileRemoval = oExchSVR.DaysBeforeLogFileRemoval
            Main.Exch.Svrs(iServers).DirectoryServer = oExchSVR.DirectoryServer
            Main.Exch.Svrs(iServers).ExchangeVersion = oExchSVR.ExchangeVersion
            Main.Exch.Svrs(iServers).MessageTrackingEnabled = oExchSVR.MessageTrackingEnabled
            Main.Exch.Svrs(iServers).SubjectLoggingEnabled = oExchSVR.SubjectLoggingEnabled
            Main.Exch.Svrs(iServers).ServerType = sServerType(oExchSVR.ServerType)
            
            ' StorageGroup Settings
            Main.Exch.Svrs(iServers).StorageGroupCount = UBound(oExchSVR.StorageGroups) + 1
            
            frmExMerge.List4.AddItem "StorageGroupCount found: '" & Main.Exch.Svrs(iServers).StorageGroupCount
            ReDim Main.Exch.Svrs(iServers).SG(Main.Exch.Svrs(iServers).StorageGroupCount)
            
            GetStorageGroups_Main Main.Exch.Svrs(iServers).name
            
            Exit Function
            iServers = iServers + 1
            
SvrSkip:
                    
            Rs.MoveNext
        Wend
        
                
        'sort Servers
        'Sort_Array Main.Exch.Svrs
        
        Rs.Close
    End If

    ' Clean up.
    Conn.Close
    Set Rs = Nothing
    Set Com = Nothing
    Set Conn = Nothing
    Exit Function
    
errHandler:
    'If UBound(Main.Exch.Svrs(iServers).Errors) Then
    '    ReDim Main.Exch.Svrs(iServers).Errors(1)
    'Else
    '    ReDim Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) + 1)
    'End If
    'Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) - 1).Number = Err.Number
    'Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) - 1).Desc = Err.Description
    'Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) - 1).Source = Err.Source
    'Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) - 1).DateTime = Now
    
    'If Err.Number = -2147024809 Then
    '    Main.Exch.Svrs(iServers).Errors(UBound(Main.Exch.Svrs(iServers).Errors) - 1).Meaning = "May be an Exchange Server Found in AD, but is on a different network or is no longer active."
    'End If
    
    DebugPrint "GetStorageServers_Main"
    frmExMerge.List4.AddItem "*GetStorageServers_Main - " & Err.Number & vbNewLine & Err.Description
    
    Resume Next
End Function

Sub DebugPrint(sMsg As String)
    Debug.Print sMsg & " - [Err: " & Err.Number & " - Desc: " & Err.Description & "]"
End Sub




Function GetStorageGroups_Main(sServerName As Variant)
    Dim oTmpSG As CDOEXM.StorageGroup
    Dim oTmpDB As CDOEXM.MailboxStoreDB
    Dim tmp As Variant
    Dim j As Integer, k As Integer
    
On Error GoTo errHandler:

    oExchSVR.DataSource.Open sServerName
    arrEXSG = oExchSVR.StorageGroups
    ReDim arrSGs(UBound(arrEXSG))
            
    For iStorageGroups = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        
    On Error GoTo NextStep:
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(iStorageGroups)
    
        Set arrSGs(iStorageGroups) = oTmpSG
        
        frmExMerge.Output_2_List frmExMerge.List4, "Digging Storage Group '" & oTmpSG.name & "'"
        
        'StorageGroup Name
        Main.Exch.Svrs(iServers).SG(iStorageGroups).name = oTmpSG.name
        Main.Exch.Svrs(iServers).SG(iStorageGroups).LogFilePath = oTmpSG.LogFilePath
        Main.Exch.Svrs(iServers).SG(iStorageGroups).SystemFilePath = oTmpSG.SystemFilePath
        Main.Exch.Svrs(iServers).SG(iStorageGroups).ZeroDatabase = oTmpSG.ZeroDatabase
        Main.Exch.Svrs(iServers).SG(iStorageGroups).CircularLogging = oTmpSG.CircularLogging
        Main.Exch.Svrs(iServers).SG(iStorageGroups).FieldCount = oTmpSG.Fields.Count
        
        
        frmExMerge.List4.AddItem "StorageGroup: '" & oTmpSG.name & "' with " & oTmpSG.Fields.Count & " fields."
        
        
        'Get MailBoxStoreDBs
        Main.Exch.Svrs(iServers).SG(iStorageGroups).MailBoxStoreDBCount = UBound(oTmpSG.MailBoxStoreDBs) + 1
        ReDim Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(Main.Exch.Svrs(iServers).SG(iStorageGroups).MailBoxStoreDBCount)
        iMailboxStoreDB = 0
        
        For Each tmp In oTmpSG.MailBoxStoreDBs
                    
            Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
            oTmpDB.DataSource.Open "LDAP://" & tmp
                
            frmExMerge.Output_2_List frmExMerge.List4, "Digging MailBoxStore Database '" & oTmpDB.name & "'"
            
            'MailBoxStoreDB Name
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).name = oTmpDB.name
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).DaysBeforeDeletedMailboxCleanup = oTmpDB.DaysBeforeDeletedMailboxCleanup
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).DaysBeforeGarbageCollection = oTmpDB.DaysBeforeGarbageCollection
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).DBPath = oTmpDB.DBPath
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).Enabled = oTmpDB.Enabled
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).GarbageCollectOnlyAfterBackup = oTmpDB.GarbageCollectOnlyAfterBackup
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).OfflineAddressList = oTmpDB.OfflineAddressList
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).OverQuotaLimit = oTmpDB.OverQuotaLimit
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).HardLimit = oTmpDB.HardLimit
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).PublicStoreDB = oTmpDB.PublicStoreDB
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).SLVPath = oTmpDB.SLVPath
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).Status = sStoreStatus(oTmpDB.Status)
            Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).StoreQuota = oTmpDB.StoreQuota
            
            frmExMerge.List4.AddItem " + MailBoxStoreDB: '" & oTmpDB.name
            
            'Grab MailBoxes + Details
            If frmExMerge.chkDiscoverOnlineDBsOnly.Value = 1 Then
                If oTmpDB.Status = cdoexmOnline Then
                    Call EnumerateFields_Main("StorageGroup fields", oTmpDB.Fields)
                End If
            Else
                Call EnumerateFields_Main("StorageGroup fields", oTmpDB.Fields)
            End If
                    
            iMailboxStoreDB = iMailboxStoreDB + 1
                
        Next
        
        'sort MailBoxStoreDBs
        'SortArray Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB
NextStep:
                            
        Set oTmpSG = Nothing
    Next
        
    'sort StorageGroups
    'SortArray Main.Exch.Svrs(iServers).SG
                    
    Exit Function
    
errHandler:
    DebugPrint "GetStorageGroups_Main"
    frmExMerge.List4.AddItem "*GetStorageGroups_Main - " & Err.Number & vbNewLine & Err.Description
    Resume Next
End Function

Function sStoreStatus(iType)
 If iType = 0 Then
  sStoreStatus = "Store is online (0)"
 ElseIf iType = 1 Then
  sStoreStatus = "Store is offline (1)"
 ElseIf iType = 2 Then
  sStoreStatus = "Store is mounting (2)"
 ElseIf iType = 3 Then
  sStoreStatus = "Store is dismounting (3)"
 Else
  sStoreStatus = "Unknown (" & iType & ")"
 End If
End Function

Function EnumerateFields_Main(str, fieldlist)
    Dim f, t
    Dim inx As Integer
        
On Error GoTo Err:
                      
    frmExMerge.Output_2_List frmExMerge.List4, "Enumerating Mailboxes "
    frmExMerge.Label50.Caption = ""
    
    iMailBoxes = 0
    inx = 0
    While inx < fieldlist.Count
        f = fieldlist(inx)
        
        'If fieldlist.Count > 0 Then
        '    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MailBoxCount = fieldlist.Count + 1
        '    ReDim Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MailBoxCount)
        'End If
        
        'MsgBox fieldlist.Count
        'Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).Name = fieldlist(iMailBoxes).Name
        
        Call HandleSingleField_Main(0, f, fieldlist(inx).name)
        inx = inx + 1
    Wend
    
    frmExMerge.Label50.Caption = ""
    'sort MailBoxes
    'SortArray Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX

    Exit Function
    
Err:
    DebugPrint "EnumerateFields_Main"
    Resume Next

End Function


Sub HandleSingleField_Main(indent, f, strName)
     Dim t
     Dim bVerbose As Boolean
    
On Error GoTo Err:

     t = TypeName(f)
     
     If t = "String" Or t = "Boolean" Or t = "Date" Or t = "Long" Or t = "Byte" Then
        
     
        If indent = 0 Then
            'Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).Alias = iMailBoxes & ".(" & strName & ", " & t & ")" & " - " & f
            'List1.AddItem "Field#" & iMailBoxes & " (" & strName & ", " & t & ") " & " " & f
        ElseIf indent > 0 Then
            'With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes)
            '    .Alias = .Alias & ", " & "(" & strName & ", " & t & ")" & " - " & f
            'End With
            'List1.AddItem Space(indent * 4) & "Index#" & iMailBoxes & " (" & strName & ", " & t & ") " & " " & f
            bVerbose = True
            If strName = "homeMDBBL" And t = "String" And bVerbose Then
                ' f contains an adspath to a mailbox user
                DumpUserInfo_Main f
            End If
            
        End If
        'List1.ListIndex = List1.ListCount - 1
        
        Exit Sub
     End If
    
    ' If Right (t, 2) = "()" Then
     If t = "Variant()" Then
        Dim i As Integer
    
        'List1.AddItem Space(indent * 4) & "Field#" & iMailBoxes & " (" & strName & ", " & t & ") "
        If strName = "homeMDBBL" Then
            With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB)
                .MailBoxCount = UBound(f) + 1
                ReDim .MBX(.MailBoxCount)
            End With
        End If
        
        For i = LBound(f) To UBound(f)
            Call HandleSingleField_Main(indent + 1, f(i), strName)
        Next
        
        'With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes)
        '    .Alias = iMailBoxes & ".(" & .Alias & ")"
        'End With
        'List1.ListIndex = List1.ListCount - 1
        
        Exit Sub
    Else
        'Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).Alias = iMailBoxes & ".(" & strName & ", " & t & ")"
    End If
    
     'List1.AddItem "Field#" & iMailBoxes & " (" & strName & ", " & t & ") "
    Exit Sub
    
Err:
    DebugPrint "HandleSingleField_Main"
    Resume Next
End Sub


Sub DumpUserInfo_Main(strADSPath)
    Dim objPerson
    Dim objRecipient As CDOEXM.IMailRecipient ' as CDOEXM.IMailRecipient
    Dim sAlias As String
    
On Error GoTo Err:
    
    Set objPerson = CreateObject("CDO.Person")
    objPerson.DataSource.Open "LDAP://" & strADSPath
    Set objRecipient = objPerson.GetInterface("IMailRecipient")
            
    With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB)
    
    
    frmExMerge.Label50.Caption = iMailBoxes & ") " & objRecipient.Alias
    frmExMerge.Label60.Caption = " of " & .MailBoxCount
    Pause 0.1
    
            
On Error GoTo errHandler:
               
        'If LCase(objRecipient.Alias) = "mschaefer" Then
        '    MsgBox objRecipient.Alias
        'End If
               
        .MBX(iMailBoxes).Alias = objRecipient.Alias
        .MBX(iMailBoxes).ForwardingStyle = objRecipient.ForwardingStyle
        .MBX(iMailBoxes).ForwardTo = objRecipient.ForwardTo
        .MBX(iMailBoxes).HideFromAddressBook = objRecipient.HideFromAddressBook
        .MBX(iMailBoxes).IncomingLimit = objRecipient.IncomingLimit
        .MBX(iMailBoxes).OutgoingLimit = objRecipient.OutgoingLimit
        .MBX(iMailBoxes).ProxyAddresses = sProxyAddresses(3, objRecipient.ProxyAddresses)
        .MBX(iMailBoxes).RestrictedAddresses = objRecipient.RestrictedAddresses
        .MBX(iMailBoxes).RestrictedAddressList = sProxyAddresses(3, objRecipient.RestrictedAddressList)
        .MBX(iMailBoxes).SMTPEmail = objRecipient.SMTPEmail
        .MBX(iMailBoxes).TargetAddress = objRecipient.TargetAddress
        
        GetSize_Count_Main Main.Exch.Svrs(iServers).name, .MBX(iMailBoxes).Alias
        LoadUser GetADUser(.MBX(iMailBoxes).Alias), .MBX(iMailBoxes).ADInfo
        
        
        frmExMerge.List4.AddItem "   ---User: '" & objRecipient.Alias
        
        'MsgBox .MBX(iMailBoxes).ADInfo.FullName
        
        'GetADUser_Info .MBX(iMailBoxes).SMTPEmail
        
errHandler:

    End With
    
    
    iMailBoxes = iMailBoxes + 1
    Set objRecipient = Nothing
    Set objPerson = Nothing

    Exit Sub
    
Err:
    DebugPrint "DumpUserInfo_Main"
    frmExMerge.List4.AddItem "*DumpUserInfo_Main - " & Err.Number & vbNewLine & Err.Description
    Resume Next
End Sub



Sub LoadUser(oUser As IADsUser, mUser As ADUser_Struct)
    On Error Resume Next
    
        'mUser.Domain = strDomain
        mUser.AccountDisabled = oUser.AccountDisabled
        mUser.AccountExpirationDate = oUser.AccountExpirationDate
        mUser.ADsPath = oUser.ADsPath
        mUser.BadLoginAddress = oUser.BadLoginAddress
        mUser.BadLoginCount = oUser.BadLoginCount
        mUser.Class = oUser.Class
        mUser.Department = oUser.Department
        mUser.Description = oUser.Description
        mUser.Division = oUser.Division
        mUser.EmailAddress = oUser.EmailAddress
        mUser.EmployeeID = oUser.EmployeeID
        mUser.FaxNumber = oUser.FaxNumber
        mUser.FirstName = oUser.FirstName
        'mUser.FullName = oUser.FullName
        mUser.GraceLoginsAllowed = oUser.GraceLoginsAllowed
        mUser.GraceLoginsRemaining = oUser.GraceLoginsRemaining
        mUser.Guid = oUser.Guid
        mUser.HomeDirectory = oUser.HomeDirectory
        mUser.HomePage = oUser.HomePage
        mUser.IsAccountLocked = oUser.IsAccountLocked
        mUser.Languages = oUser.Languages
        mUser.LastFailedLogin = oUser.LastFailedLogin
        mUser.LastLogin = oUser.LastLogin
        mUser.LastLogoff = oUser.LastLogoff
        mUser.LastName = oUser.LastName
        mUser.LoginHours = oUser.LoginHours
        mUser.LoginScript = oUser.LoginScript
        mUser.LoginWorkstations = oUser.LoginWorkstations
        mUser.Manager = oUser.Manager
        mUser.MaxLogins = oUser.MaxLogins
        mUser.MaxStorage = oUser.MaxStorage
        mUser.name = oUser.name
        mUser.NamePrefix = oUser.NamePrefix
        mUser.NameSuffix = oUser.NameSuffix
        mUser.OfficeLocations = oUser.OfficeLocations
        mUser.OtherName = oUser.OtherName
        mUser.Parent = oUser.Parent
        mUser.PasswordExpirationDate = oUser.PasswordExpirationDate
        mUser.PasswordLastChanged = oUser.PasswordLastChanged
        mUser.PasswordMinimumLength = oUser.PasswordMinimumLength
        mUser.PasswordRequired = oUser.PasswordRequired
        mUser.Picture = oUser.Picture
        mUser.PostalAddresses = oUser.PostalAddresses
        mUser.PostalCodes = oUser.PostalCodes
        mUser.Profile = oUser.Profile
        mUser.RequireUniquePassword = oUser.RequireUniquePassword
        mUser.Schema = oUser.Schema
        mUser.SeeAlso = oUser.SeeAlso
        mUser.TelephoneHome = oUser.TelephoneHome
        mUser.TelephoneMobile = oUser.TelephoneMobile
        mUser.TelephoneNumber = oUser.TelephoneNumber
        mUser.TelephonePager = oUser.TelephonePager
        mUser.Title = oUser.Title
        
End Sub





Sub GetADUser_Info(sEmail As String)
    Dim objComputer As IADsMembers
    Dim objUser As IADsUser
    Dim Param
    Dim i As Integer
    
    Param = Split(sEmail, "@")
    With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB)
    
    If UBound(Param) > 0 Then
    
        Set objComputer = GetObject("WinNT://" & Param(0))
        objComputer.Filter = Array("User")
        
        i = 0
        For Each objUser In objComputer
            If objUser.EmailAddress = sEmail Then
                .MBX(iMailBoxes).FullName = objUser.FullName
                .MBX(iMailBoxes).name = objUser.name
            End If
        Next
        
    End If
    
    End With
    
End Sub

Function sProxyAddresses(iTabs, objProxyAddresses)
    Dim strHdr, strOut, strProxy
    Dim i
    
    strOut = ""
    strHdr = ""
    
    For i = 0 To (iTabs - 1)
        strHdr = strHdr & vbTab
    Next
    
    i = 0
    For Each strProxy In objProxyAddresses
        If i = 0 Then
            strOut = strProxy
        Else
            strOut = strOut & vbCrLf & strHdr & strProxy
        End If
        
        i = i + 1
    Next
    
    sProxyAddresses = strOut
End Function

Function sServerType(iType)
    If iType = 0 Then
        sServerType = "Backend (0)"
    ElseIf iType = 1 Then
        sServerType = "Frontend (1)"
    Else
        sServerType = "Unknown (" & iType & ")"
    End If
End Function

Function GetSize_Count_Main(ServerName As String, MailBox As String) As String
   Dim oSession
   Dim oInfoStores
   Dim oInfoStore
   Dim StorageUsed
   Dim StorageUsedMega
   Dim NumMessages
   Dim strProfileInfo
   Dim sMsg
   Dim name() As String
   Const OneKiloByte = 1024
   Const OneMegaByte = 1048576 '1048576
   'Const OneGigaByte = 1024 * 1024 * 1024 '1048576
   
    Const PR_NTSDModificationTime = &H3FD60040
    Const PR_Creation_Time = &H30070040
   

    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).MessageCount = -1
    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).Size = 0
    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).DisplaySize = -1
    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).DateCreated = ""
    Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).DateLastModified = ""
    
On Error GoTo Err:

   'Create Session object.
   Set oSession = CreateObject("MAPI.Session")
   If Err.Number <> 0 Then
        DebugPrint "GetSize_Count_Main:Error creating MAPI.Session"
      'sMsg = "Error creating MAPI.Session."
      'sMsg = sMsg & "Make sure CDO 1.21 is installed. "
      'sMsg = sMsg & Err.Number & " " & Err.Description
      'Debug.Print sMsg
      Exit Function
   End If
    
   strProfileInfo = ServerName & vbLf & MailBox

   'Log on.
   oSession.Logon , , False, True, , True, strProfileInfo
   If Err.Number <> 0 Then
        DebugPrint "GetSize_Count_Main:Error logging on"
      'sMsg = "Error logging on: "
      'sMsg = sMsg & Err.Number & " " & Err.Description
      'Debug.Print sMsg
      'Debug.Print "Server: " & ServerName
      'Debug.Print "Mailbox: " & MailBox
      Set oSession = Nothing
      Exit Function
   End If

    'Grab the information stores.
    Set oInfoStores = oSession.InfoStores
    If Err.Number <> 0 Then
        DebugPrint "GetSize_Count_Main:Error retrieving InfoStores Collection"
        'sMsg = "Error retrieving InfoStores Collection: "
        'sMsg = sMsg & Err.Number & " " & Err.Description
        'Debug.Print sMsg
        'Debug.Print "Server: " & ServerName
        'Debug.Print "Mailbox: " & MailBox
        Set oInfoStores = Nothing
        Set oSession = Nothing
        Exit Function
    End If
    
    
    'Loop through information stores to find the user's mailbox.
    For Each oInfoStore In oInfoStores
        'MsgBox oInfoStore.name
        If InStr(1, oInfoStore.name, "Mailbox - ", 1) <> 0 Then
            name = Split(oInfoStore.name, " - ")
            If UBound(name) > 0 Then
                Debug.Print name(1)
                If LCase(name(1)) = "meg schaefer" Then
                    Debug.Print name(1)
                End If
                Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).ADInfo.FullName = name(1)
            End If
            
            'Dim oField
            'Set oField = CreateObject("MAPI.Session")
            'Dim lMbxSizeKB  As Long

            'oField = oInfoStore.Fields(&HE080003)
            'MsgBox oField.value
            'lMbxSizeKB = CLng(oField.value)
            'lMbxSizeKB = CLng(oField.value / OneKiloByte)
            
            '&HE080003 = PR_MESSAGE_SIZE
            
            'StorageUsed = CLng(oInfoStore.Fields(&HE080003)) / OneMegaByte
            StorageUsed = CLng(oInfoStore.Fields(&HE080014)) / OneMegaByte
            StorageUsedMega = CLng(oInfoStore.Fields(&HE080014)) / OneMegaByte
            
            'Debug.Print oInfoStore.Fields(&H30080040)
            
            'Dim dblStorageUsed As Double
            
            
'On Error GoTo spit:

            'If Sgn(StorageUsed) = -1 Then                   ' is it a negative number
            '        dblStorageUsed = 2147483647                     ' yes, so it's really > 2GB
            '        StorageUsed = StorageUsed * -1          ' reverse the sign
            '        dblStorageUsed = dblStorageUsed + StorageUsed   ' add the result
            'Else
            '        dblStorageUsed = CDbl(StorageUsed)      ' just use the value as-is
            'End If
            
            'Debug.Print Int(dblStorageUsed / 1048576)
            
            'Dim pi As Integer
            'Dim v
            
            
            'For pi = 0 To 64
            '    v = &HE080000 + pi
            '    'Debug.Print &HE080003
            '    Debug.Print v & ":" & oInfoStore.Fields(v)
            'Next pi
            'Debug.Print &HE080014
            
'spit:
            
            
            If Err.Number <> 0 Then
                DebugPrint "GetSize_Count_Main:Error retrieving PR_MESSAGE_SIZE"
               'sMsg = "Error retrieving PR_MESSAGE_SIZE: "
               'sMsg = sMsg & Err.Number & " " & Err.Description
               'Debug.Print sMsg
               'Debug.Print "Server: " & ServerName
               'Debug.Print "Mailbox: " & MailBox
               Set oInfoStore = Nothing
               Set oInfoStores = Nothing
               Set oSession = Nothing
               Exit Function
            End If
         
         '&H33020003 = PR_CONTENT_COUNT
         NumMessages = oInfoStore.Fields(&H36020003)

         If Err.Number <> 0 Then
            DebugPrint "GetSize_Count_Main:Error Retrieving PR_CONTENT_COUNT"
            'sMsg = "Error Retrieving PR_CONTENT_COUNT: "
            'sMsg = sMsg & Err.Number & " " & Err.Description
            'Debug.Print sMsg
            'Debug.Print "Server: " & ServerName
            'Debug.Print "Mailbox: " & MailBox
            Set oInfoStore = Nothing
            Set oInfoStores = Nothing
            Set oSession = Nothing
            Exit Function
         End If

        'GetSize_Count = vbTab & "| " & NumMessages & "| " & StorageUsed
        Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).MessageCount = FormatNumber(NumMessages, 0, vbTrue, vbTrue, vbTrue)
        Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).Size = StorageUsed
        Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB).MBX(iMailBoxes).DisplaySize = FormatNumber(StorageUsedMega, 0, vbTrue, vbTrue, vbTrue)
         'sMsg = "Storage Used in " & oInfoStore.name
         'sMsg = sMsg & " (bytes): " & StorageUsed
         'Debug.Print sMsg
         'Debug.Print "Number of Messages: " & NumMessages
        End If
    Next
   
    ' attempt to grab the DateCreated and DateLastModified fields
On Error GoTo skipDate:

    Dim objInbox, objInfoStore, objRootfolder, Non_IPM_rootFolder

    Set objInbox = oSession.Inbox
    Set objInfoStore = oSession.GetInfoStore(objInbox.StoreID)
    Set objRootfolder = objInfoStore.RootFolder
    
    With Main.Exch.Svrs(iServers).SG(iStorageGroups).MBSDB(iMailboxStoreDB)
        Set Non_IPM_rootFolder = oSession.GetFolder(objRootfolder.Fields.Item(&HE090102), objInfoStore.ID)
        'List1.AddItem Non_IPM_rootFolder.Fields.Item(PR_NTSDModificationTime)
        .MBX(iMailBoxes).DateLastModified = Format(Non_IPM_rootFolder.Fields.Item(PR_NTSDModificationTime), "mm/dd/yyyy")
           
           
        Set Non_IPM_rootFolder = oSession.GetFolder(objRootfolder.Fields.Item(&HE090102), objInfoStore.ID)
        'List1.AddItem Non_IPM_rootFolder.Fields.Item(PR_Creation_Time)
        .MBX(iMailBoxes).DateCreated = Format(Non_IPM_rootFolder.Fields.Item(PR_Creation_Time), "mm/dd/yyyy")
    End With

skipDate:

    ' Log off.
    oSession.Logoff

    ' Clean up memory.
    Set oInfoStore = Nothing
    Set oInfoStores = Nothing
    Set oSession = Nothing
    Exit Function
    
Err:
    DebugPrint "GetSize_Count_Main:Err"
    Resume Next
    Set oInfoStore = Nothing
    Set oInfoStores = Nothing
    Set oSession = Nothing
    Exit Function
End Function





Function GetADUser(sName As String) As IADsUser
    Dim iAdRootDSE As IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim Filter As Variant
    Dim varConfigNC As Variant
    Dim strQuery As String
    
    Dim oUser As IADsUser
        
On Error GoTo errHandler:

    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    varConfigNC = iAdRootDSE.Get("defaultNamingContext")

    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"
    
    strQuery = "<LDAP://" & varConfigNC & ">;(mailnickname=" & sName & ");samaccountname,distinguishedName;subtree"

    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        Set GetADUser = GetObject("LDAP://" & Rs.Fields("distinguishedName").Value)
    End If
    
    Exit Function
    
errHandler:
    Resume Next
End Function



