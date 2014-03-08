Attribute VB_Name = "mod_CDOEXM_Exchange"

'Dim oExchangeServer As New CDOexm.ExchangeServer
'Dim arrSGs() As CDOexm.StorageGroup

'Dim oExchangeDB As New CDOexm.MailBoxStoreDB
'Dim arrMSs() As CDOexm.MailBoxStoreDB

Sub GetServerDetails(ServerName As String)

On Error GoTo errHandler:

    'Try to connect
    oExchangeServer.DataSource.Open ServerName
    
    'List1.AddItem "1." & oExchangeServer.DaysBeforeLogFileRemoval
    'List1.AddItem "2." & oExchangeServer.DirectoryServer
    'List1.AddItem "3." & oExchangeServer.ExchangeVersion
    'List1.AddItem "4." & oExchangeServer.MessageTrackingEnabled
    'List1.AddItem "5." & oExchangeServer.SubjectLoggingEnabled
    'List1.AddItem "6." & UBound(oExchangeServer.StorageGroups) + 1
    
    If oExchangeServer.ServerType = cdoexmBackEnd Then
        'List1.AddItem "7." & "Backend"
    Else
        'List1.AddItem "7." & "Frontend"
    End If

    'Scroll through the array of storage groups and
    'open each one into a SG object
    arrEXSG = oExchangeServer.StorageGroups
    Dim oTmpSG
    
    Dim oTmpDB
    Dim odb As Object
    
    Dim tmp As Variant
    
    ReDim arrSGs(UBound(arrEXSG))
    
    'grab StorageGroups per ExchangeServer
    For i = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(i)
        Set arrSGs(i) = oTmpSG
        'Fill in the combo box
        'List1.AddItem oTmpSG.name
        'Debug.Print oTmpSG.Name
    
        'grab MailBoxStoreDBs per StorageGroup
        For Each tmp In oTmpSG.MailBoxStoreDBs
            Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
            oTmpDB.DataSource.Open "LDAP://" & tmp
            'List1.AddItem "  -->" & oTmpDB.name
        Next
    
        Set oTmpSG = Nothing
    Next
    
    'List1.ListIndex = 0
    
    Exit Sub
errHandler:
    MsgBox "There was an error in cmdConnect_Click(). Error#" _
         & Err.Number & " Description:" & Err.Description, _
           vbCritical + vbOKOnly
    End
End Sub

Private Sub connect()
    On Error GoTo errHandler:
    If txtExServerName.Text = "" Then
        MsgBox "You must enter a name in the Exchange Server name box!", _
               vbCritical + vbOKOnly
        Exit Sub
    End If
    
    If Right(txtExServerName.Text, 7) = "LDAP://" Then
        'Remove the LDAP
        txtExServerName.Text = Mid(txtExServerName.Text, 8)
    End If
    
    'Try to connect
    oExchangeServer.DataSource.Open txtExServerName.Text
    
    'Try to fill in Exchange Server info
    FillInExchangeServerInfo
    
    'Retrieve the storage groups
    FillInStorageGroups
    
    Exit Sub
errHandler:
    MsgBox "There was an error in cmdConnect_Click(). Error#" _
         & Err.Number & " Description:" & Err.Description, _
           vbCritical + vbOKOnly
    End
End Sub
    
Sub FillInExchangeServerInfo()
    On Error GoTo errHandler
    'List1.AddItem "1." & oExchangeServer.DaysBeforeLogFileRemoval
    'List1.AddItem "2." & oExchangeServer.DirectoryServer
    'List1.AddItem "3." & oExchangeServer.ExchangeVersion
    'List1.AddItem "4." & oExchangeServer.MessageTrackingEnabled
    'List1.AddItem "5." & oExchangeServer.SubjectLoggingEnabled
    'List1.AddItem "6." & UBound(oExchangeServer.StorageGroups) + 1
    
    If oExchangeServer.ServerType = cdoexmBackEnd Then
        'List1.AddItem "Backend"
    Else
        'List1.AddItem "Frontend"
    End If
    
    Exit Sub
errHandler:
    MsgBox "There was an error in FillInExchangeServerInfo. Error#" _
         & Err.Number & " Description:" & Err.Description, _
         vbCritical + vbOKOnly
End Sub
    
Sub FillInExchangeServerInfo2()
    On Error GoTo errHandler
    lblLogRemoval.Caption = oExchangeServer.DaysBeforeLogFileRemoval
    lblDirServer.Caption = oExchangeServer.DirectoryServer
    lblVersion.Caption = oExchangeServer.ExchangeVersion
    lblMsgTracking.Caption = oExchangeServer.MessageTrackingEnabled
    lblSubjectLogging.Caption = oExchangeServer.SubjectLoggingEnabled
    lblSGCount.Caption = UBound(oExchangeServer.StorageGroups) + 1
    If oExchangeServer.ServerType = cdoexmBackEnd Then
        lblServerType.Caption = "Backend"
    Else
        lblServerType.Caption = "Frontend"
    End If
    
    Exit Sub
errHandler:
    MsgBox "There was an error in FillInExchangeServerInfo. Error#" _
         & Err.Number & " Description:" & Err.Description, _
         vbCritical + vbOKOnly
End Sub
Sub FillInStorageGroups()
    On Error GoTo errHandler
    
    'Scroll through the array of storage groups and
    'open each one into a SG object
    arrEXSG = oExchangeServer.StorageGroups
    Dim oTmpSG
    
    Dim oTmpDB
    Dim odb As Object
    
    Dim tmp As Variant
    
    ReDim arrSGs(UBound(arrEXSG))
    For i = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(i)
        Set arrSGs(i) = oTmpSG
        'Fill in the combo box
        List1.AddItem oTmpSG.name
        'Debug.Print oTmpSG.Name
    
        For Each tmp In oTmpSG.MailBoxStoreDBs
            Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
            oTmpDB.DataSource.Open "LDAP://" & tmp
            'List1.AddItem "  -->" & oTmpDB.name
        Next
    
        Set oTmpSG = Nothing
    Next
    
    'List1.ListIndex = 0
    
    Exit Sub
errHandler:
    MsgBox "There was an error in FillInStorageGroups. Error#" & Err.Number _
         & " Description:" & Err.Description, vbCritical + vbOKOnly
End Sub

Sub getStorageDBs()
    Dim theServer       'As CDOEXM.ExchangeServer
    Dim theMDB          'As CDOEXM.MailboxStoreDB
    Dim strServerName As Variant
    Dim strMDBName As Variant
    Dim theFirstSG As Variant
    Dim db As Object
    
    strServerName = "susamail"
    strMDBName = "Mailbox Store (BATMAN)"
    
    Set theServer = CreateObject("CDOEXM.ExchangeServer")
    Set theMDB = CreateObject("CDOEXM.MailboxStoreDB")
    
    theServer.DataSource.Open strServerName
    
    For Each SG In theServer.StorageGroups
        theFirstSG = SG
        Exit For
    Next
    
    strURL = "LDAP://" & theServer.DirectoryServer & "/cn=" & strMDBName & "," & theFirstSG
    theMDB.DataSource.Open strURL
    
    
    'theMDB.Mount
End Sub


Sub FillInStorageDatabases(SG As Variant)
    On Error GoTo errHandler
    
    'Scroll through the array of storage groups and
    'open each one into a SG object
    'oExchangeDB.DataSource.Open SG.MailboxStoreDBs
    
    
    'Dim imbxDB As CDOEXM.MailboxStoreDB
    'Dim iSG As CDOEXM.StorageGroup
    'Dim mbx As Variant
    Dim vsg As Variant
    
    For Each vsg In SG
        MsgBox vsg.name
    Next
    
    'iSG.DataSource.Open SG
    'For Each mbx In iSG.MailboxStoreDBs
        'List1.AddItem imbxDB.Name
    'Next
    
    'vbscript
    'Dim storegroup    ' as Variant (string)
    'Dim imbxDB        ' as CDOEXM.MailBoxStoreDB
    'Dim mbx           ' as Variant (string)
    
    'Dim isg           ' as CDOEXM.StorageGroup
    'Set isg = CreateObject("CDOEXM.StorageGroup")
    'isg.DataSource.Open storegroup
    
    'For Each mbx In isg.MailboxStoreDBs
    
        'imbxDB.DataSource.Open mbx
        'e "Name                            = " & imbxDB.Name
        'e "DaysBeforeDeletedMailboxCleanup = " & imbxDB.DaysBeforeDeletedMailboxCleanup
        'e "DaysBeforeGarbageCollection     = " & imbxDB.DaysBeforeGarbageCollection
        'e "DBPath                          = " & imbxDB.DBPath
        'e "Enabled                         = " & imbxDB.Enabled
        'e "GarbageCollectOnlyAfterBackup   = " & imbxDB.GarbageCollectOnlyAfterBackup
        'e "OfflineAddressList              = " & imbxDB.OfflineAddressList
        'e "OverQuotaLimit                  = " & imbxDB.OverQuotaLimit
        'e "HardLimit                       = " & imbxDB.HardLimit
        'e "PublicStoreDB                   = " & imbxDB.PublicStoreDB
        'e "SLVPath                         = " & imbxDB.SLVPath
        'e "Status                          = " & sStoreStatus(imbxDB.Status)
        'e "StoreQuota                      = " & imbxDB.StoreQuota
        'Call EnumerateFields("Mailbox Store fields", imbxDB.Fields)
        
    'Next

    
    
    
    'Dim oTmpMS As CDOEXM.IMailboxStore
    
    'List1.ListIndex = 0
    
    Exit Sub
errHandler:
    MsgBox "There was an error in FillInStorageGroups. Error#" & Err.Number _
         & " Description:" & Err.Description, vbCritical + vbOKOnly
End Sub




Sub FillInStorageGroups2()
    On Error GoTo errHandler
    
    'Scroll through the array of storage groups and
    'open each one into a SG object
    arrEXSG = oExchangeServer.StorageGroups
    Dim oTmpSG 'As CDOEXM.StorageGroup
    ReDim arrSGs(UBound(arrEXSG))
    For i = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(i)
        Set arrSGs(i) = oTmpSG
        'Fill in the combo box
        comboSG.AddItem oTmpSG.name
        Set oTmpSG = Nothing
    Next
    
    comboSG.ListIndex = 0
    
    Exit Sub
errHandler:
    MsgBox "There was an error in FillInStorageGroups. Error#" & Err.Number _
         & " Description:" & Err.Description, vbCritical + vbOKOnly
End Sub
    
Sub FillInSGInfo(iIndex)
    lblCL.Caption = arrSGs(iIndex).CircularLogging
    lblLFPath.Caption = arrSGs(iIndex).LogFilePath
    lblMBDB.Caption = UBound(arrSGs(iIndex).MailBoxStoreDBs) + 1
    lblPFDB.Caption = UBound(arrSGs(iIndex).PublicStoreDBs) + 1
    lblZeroDB.Caption = arrSGs(iIndex).ZeroDatabase
    lblSFPath.Caption = arrSGs(iIndex).SystemFilePath
End Sub

Sub temp()
    Dim oMBd As New CDOEXM.MailboxStoreDB
    'MsgBox oMBd
End Sub


Sub CreateNewDatabaseStore()
    Dim oMB 'As New CDOEXM.MailboxStoreDB
    Set oMB = CreateObject("CDOEXM.MailboxStoreDB")
    'oMB.DataSource.SaveTo "LDAP://SERVER/CN=MyNewMailboxDB," _
        '& "CN=First Storage Group,CN=InformationStore,CN=SERVER," _
        '& "CN=Servers,CN=First Administrative Group," _
        '& "CN=Administrative Groups,CN=First Organization," _
        '& "CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=DOMAIN,DC=com"
        
    'oMB.Mount
End Sub

Private Sub cmdConnect_Click()
    connect
End Sub





