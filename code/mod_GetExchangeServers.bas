Attribute VB_Name = "mod_GetExchangeServers"

'Private oExchSVR As New CDOexm.ExchangeServer
'Private arrSGs() As CDOexm.StorageGroup

'Private oExchangeDB As New CDOexm.MailBoxStoreDB
'Private arrMSs() As CDOexm.MailBoxStoreDB

'Private oMBX As CDOexm.IExchangeMailbox

Private sCurrentServer As String


' Enumerate Exchange Server computers with ADSI
' The following sample queries Active Directory using ADO for all of the Exchange Server
' computers in the Forest.
' This code can be run from any Windows Server or DSCLient computer.

'Private oExchSVR As New CDOEXM.ExchangeServer
Private oExchSVR As New CDOEXM.ExchangeServer



Function GetExchangeServer_Info(sServerName As String, iIndex As Integer)
            
    Main.Exch.ServerCount = 1
    ReDim Main.Exch.Svrs(1)
    Main.Exch.Svrs(iIndex).name = sServerName
    iServers = iIndex
    
On Error GoTo errHandler:

    'Set oExchSVR = GetObject("CDOEXM.ExchangeServer")

    oExchSVR.DataSource.Open sServerName
    
    Main.Exch.Svrs(iIndex).name = oExchSVR.name
    Main.Exch.Svrs(iIndex).DaysBeforeLogFileRemoval = oExchSVR.DaysBeforeLogFileRemoval
    Main.Exch.Svrs(iIndex).DirectoryServer = oExchSVR.DirectoryServer
    Main.Exch.Svrs(iIndex).ExchangeVersion = oExchSVR.ExchangeVersion
    Main.Exch.Svrs(iIndex).MessageTrackingEnabled = oExchSVR.MessageTrackingEnabled
    Main.Exch.Svrs(iIndex).SubjectLoggingEnabled = oExchSVR.SubjectLoggingEnabled
    Main.Exch.Svrs(iIndex).ServerType = sServerType(oExchSVR.ServerType)
    Main.Exch.Svrs(iIndex).StorageGroupCount = UBound(oExchSVR.StorageGroups) + 1
    ReDim Main.Exch.Svrs(iIndex).SG(Main.Exch.Svrs(iIndex).StorageGroupCount)
            
    GetStorageGroups_Main Main.Exch.Svrs(iIndex).name
        
    Exit Function
    
errHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
    Resume Next
End Function

Function GetExchangeServers()
    Dim iAdRootDSE As IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim varConfigNC As Variant
    Dim strQuery As String
    Dim varVersion() As Variant
    
On Error GoTo errHandler:

    ' Get the configuration naming context.
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    varConfigNC = iAdRootDSE.Get("configurationNamingContext")

    ''' Open the connection.
    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"

    ''' Build the query to find all Exchange Server computers.
    strQuery = "<LDAP://" & varConfigNC & ">;(objectCategory=msExchExchangeServer);name,serialNumber;subtree"

    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    If Not Rs.EOF Then
        Dim i As Integer, j As Integer, k As Integer, sTmpStore As String
        'ReDim Main.Servers(Rs.RecordCount)
    
    
        While Not Rs.EOF
            'Main.Servers(i).Version = Rs.Fields("serialNumber").Value
            'Main.Servers(i).name = Rs.Fields("name")
            
            'frmExMerge.List6.AddItem "Svr: " & Rs.Fields("name")
            sCurrentServer = Rs.Fields("name")
        
On Error GoTo NextServer:
        
            ''' Try to connect
            oExchSVR.DataSource.Open Rs.Fields("name")
            
            'Main.Servers(i).DaysBeforeLogFileRemoval = oExchSVR.DaysBeforeLogFileRemoval
            'Main.Servers(i).DirectoryServer = oExchSVR.DirectoryServer
            '''''''''Main.Servers(i).Version = oExchSVR.ExchangeVersion
            'Main.Servers(i).MessageTrackingEnabled = oExchSVR.MessageTrackingEnabled
            'Main.Servers(i).SubjectLoggingEnabled = oExchSVR.SubjectLoggingEnabled
            
            'Main.Servers(i).StorageGroupCount = UBound(oExchSVR.StorageGroups) + 1
            
            'ReDim Main.Servers(i).StorageGroups(UBound(oExchSVR.StorageGroups) + 1)
            
            'If oExchSVR.ServerType = cdoexmBackEnd Then
                'Main.Servers(i).ServerType = "7." & "Backend"
            'Else
                'Main.Servers(i).ServerType = "7." & "Frontend"
            'End If
        
            ''' Scroll through the array of storage groups and
            ''' open each one into a SG object
            arrEXSG = oExchSVR.StorageGroups
            Dim oTmpSG          'As CDOEXM.StorageGroup
            
            Dim oTmpDB          'As CDOEXM.MailboxStoreDB
            Dim odb As Object
            
            Dim tmp As Variant
            
            
            ReDim arrSGs(UBound(arrEXSG))
            
            ''' grab StorageGroups per ExchangeServer
            For j = LBound(arrEXSG) To UBound(arrEXSG)
                Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
                oTmpSG.DataSource.Open "LDAP://" & arrEXSG(j)
                Set arrSGs(j) = oTmpSG
                
                'Main.Servers(i).StorageGroups(j).name = oTmpSG.name
                'frmExMerge.List6.AddItem "   ->SG: " & oTmpSG.name
                
                
                ''' grab MailBoxStoreDBs per StorageGroup
                k = 0
                'Main.Servers(i).StorageGroups(j).MailBoxStoreDBCount = 0
                For Each tmp In oTmpSG.MailBoxStoreDBs
                
                    Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
                    oTmpDB.DataSource.Open "LDAP://" & tmp
                    
                    'ReDim Main.Servers(i).StorageGroups(j).MailBoxStoreDB(k + 1)
                    '
                    'Main.Servers(i).StorageGroups(j).MailBoxStoreDB(k).name = oTmpDB.name
                    'frmExMerge.List6.AddItem "      -->MBS: " & oTmpDB.name
                    
                    ''''''add users here
                    
                    k = k + 1
                    'Main.Servers(i).StorageGroups(j).MailBoxStoreDBCount = k
                    
                Next
                            
                Set oTmpSG = Nothing
            Next
            
NextServer:
        
            Rs.MoveNext
            i = i + 1
        Wend
        
        Rs.Close
    End If

    ' Clean up.
    Conn.Close
    Set Rs = Nothing
    Set Com = Nothing
    Set Conn = Nothing
    Exit Function
    
errHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
    Resume Next
End Function


Function GetExchangeServers_Only()
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
        Dim i As Integer
        
        ReDim Servers(Rs.RecordCount)
    
        While Not Rs.EOF
            'If i = 0 Then
            '    cbo.Text = Rs.Fields("name")
            '    sCurrentServer = Rs.Fields("name")
            'End If
        
            Servers(i).name = LCase(Trim(Rs.Fields("name")))
            'cbo.AddItem Rs.Fields("name")
        
            Rs.MoveNext
            i = i + 1
        Wend
        
        Rs.Close
    End If

    ' Clean up.
    Conn.Close
    Set Rs = Nothing
    Set Com = Nothing
    Set Conn = Nothing
    Exit Function
    
errHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Function



Function GetStorageGroups(sServerName As Variant, cbo As ComboBox)
    Dim oTmpSG      'As CDOEXM.StorageGroup
    Dim j As Integer
    
On Error GoTo errHandler:

    ''' Try to connect
    oExchSVR.DataSource.Open sServerName
            
    sCurrentServer = sServerName
    
    arrEXSG = oExchSVR.StorageGroups
            
    ReDim arrSGs(UBound(arrEXSG))
            
    For j = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(j)
        Set arrSGs(j) = oTmpSG
                
        If j = 0 Then cbo.Text = oTmpSG.name
                
        cbo.AddItem oTmpSG.name
                            
        Set oTmpSG = Nothing
    Next
            
    Exit Function
    
errHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
    Resume Next
End Function

Function GetMailBoxes22()
     'Dim objPerson
     'Dim objRecipient ' as CDOEXM.IMailRecipient
    
     'Set objPerson = CreateObject("CDO.Person")
     'objPerson.DataSource.Open "LDAP://" & strADSPath
    
     'Set objRecipient = objPerson.GetInterface("IMailRecipient")
    
     'e vbTab & vbTab & "Alias                 = " & objRecipient.Alias
     'e vbTab & vbTab & "ForwardingStyle       = " & objRecipient.ForwardingStyle
     'e vbTab & vbTab & "ForwardTo             = " & objRecipient.ForwardTo
     'e vbTab & vbTab & "HideFromAddressBook   = " & objRecipient.HideFromAddressBook
     'e vbTab & vbTab & "IncomingLimit         = " & objRecipient.IncomingLimit
     'e vbTab & vbTab & "OutgoingLimit         = " & objRecipient.OutgoingLimit
     'e  vbTab & vbTab & "ProxyAddresses        = " & sProxyAddresses(3, objRecipient.ProxyAddresses)
     'e vbTab & vbTab & "RestrictedAddresses   = " & objRecipient.RestrictedAddresses
     'e vbTab & vbTab & "RestrictedAddressList = " & sProxyAddresses(3, objRecipient.RestrictedAddressList)
     'e vbTab & vbTab & "SMTPEmail             = " & objRecipient.SMTPEmail
     'e vbTab & vbTab & "TargetAddress         = " & objRecipient.TargetAddress
    
     'Set objRecipient = Nothing
     'Set objPerson = Nothing
End Function


Function GetMailBoxStores(sServerName As Variant, sStorageGroup As Variant, cbo As ComboBox)
    Dim oTmpSG          'As CDOEXM.StorageGroup
    Dim k As Integer
                
    Dim oTmpDB          'As CDOEXM.MailboxStoreDB
    Dim tmp As Variant
                
    
On Error GoTo errHandler:

    ''' Try to connect
    oExchSVR.DataSource.Open sServerName
            
    arrEXSG = oExchSVR.StorageGroups
            
    ReDim arrSGs(UBound(arrEXSG))
            
    For j = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(j)
        Set arrSGs(j) = oTmpSG
        
        If LCase(oTmpSG.name) = LCase(Trim(sStorageGroup)) Then
        
            For Each tmp In oTmpSG.MailBoxStoreDBs
                    
                Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
                oTmpDB.DataSource.Open "LDAP://" & tmp
                
                If k = 0 Then cbo.Text = oTmpDB.name
                cbo.AddItem oTmpDB.name
                
                k = k + 1
                
            Next
            
            If k = 0 Then Exit Function
        End If
                            
        Set oTmpSG = Nothing
    Next
            
    Exit Function
    
errHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
    Resume Next
End Function



Function GetMailBoxes(sServerName As Variant, sStorageGroup As Variant, sMailBoxStore As Variant, cbo As ComboBox)
    Dim oTmpSG          'As CDOEXM.StorageGroup
    Dim k As Integer
                
    Dim oTmpDB          'As CDOEXM.MailboxStoreDB
    Dim tmp As Variant
                
    'Dim imbxDB As New CDOEXM
    
On Error GoTo errHandler:

    ''' Try to connect
    oExchSVR.DataSource.Open sServerName
            
    arrEXSG = oExchSVR.StorageGroups
            
    ReDim arrSGs(UBound(arrEXSG))
            
    For j = LBound(arrEXSG) To UBound(arrEXSG)
        Set oTmpSG = CreateObject("CDOEXM.StorageGroup")
        oTmpSG.DataSource.Open "LDAP://" & arrEXSG(j)
        Set arrSGs(j) = oTmpSG
        
        If LCase(oTmpSG.name) = LCase(Trim(sStorageGroup)) Then
        
            For Each tmp In oTmpSG.MailBoxStoreDBs
                    
                Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
                oTmpDB.DataSource.Open "LDAP://" & tmp
                
                                'MsgBox oTmpDB.name & " / " & sMailBoxStore
                
                If LCase(oTmpDB.name) = LCase(Trim(sMailBoxStore)) Then
                
                    frmExMerge.List7.Clear
                    Call EnumerateFields("StorageGroup fields", oTmpDB.Fields)
    
                    'Dim objPerson As New CDO.Person
                    'Dim stmp9  As String
                    'stmp9 = "LDAP://susamail/CN=jamessmith,CN=users,DC=CompanyADomain,DC=Dev,DC=com"
                    'stmp9 = "LDAP://susany0/CN=RecipientName,CN=container,CN=container2,CN=containerN," & _
                                        "DC=siboneyusa,DC=com"
                    'RecipientName
                    'objPerson.DataSource.Open

                    'oMBX.HomeMDB = oTmpDB.name
                    'For Each tmp In oTmpDB
                            
                       ' Set oTmpDB = CreateObject("CDOEXM.MailboxStoreDB")
                        'oTmpDB.DataSource.Open "LDAP://" & tmp
    
    
                    'Dim iAdRootDSE As IADs
                    'Dim Conn As New ADODB.Connection
                    'Dim Com As New ADODB.Command
                    'Dim Rs As ADODB.Recordset
                    'Dim varConfigNC As Variant
                    'Dim strQuery As String
                
                    ' Get the configuration naming context.
                    'Set iAdRootDSE = GetObject("LDAP://RootDSE")
                    'varConfigNC = iAdRootDSE.Get("configurationNamingContext")

                    ''' Open the connection.
                    'Conn.Provider = "ADsDSOObject"
                    'Conn.Open "ADs Provider"
                
                    ''' Build the query to find all Exchange Server computers.
                    '"<LDAP://batman>;(homeMDB=" & mailDB & ");cn;subtree"
                              'Dim GALQueryFilter, tmp2
                              
                    'GALQueryFilter = "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE) " & _
                            " (| (&(objectCategory=person)(objectClass=user)" & _
                            " (msExchHomeServerName=/o=SIBONEYUSA CORP/ou=First Administrative Group/" & _
                            "cn=Configuration/cn=Servers/cn=susamail)) )))))"
                            
                    'GALQueryFilter = "(&(&(&(& (mailnickname=*) " & _
                    '        " (| (&(objectCategory=person)(objectClass=user)" & _
                    '        " (msExchHomeServerName=/o=SIBONEYUSA CORP/ou=First Administrative Group/" & _
                    '        "cn=Configuration/cn=Servers/cn=susamail)) )))))"
                            
                    'tmp2 = "(objectCategory=msExchExchangeServer)"
                    
                    'strQuery = "<LDAP://" & varConfigNC & ">;(homeMDB=" & LCase(oTmpDB.name) & ");cn;RECIPIENTS;subtree"
                    'strQuery = "<LDAP://dc=siboneyusa,dc=com>;(homeMDB=" & LCase(oTmpDB.name) & ");" & GALQueryFilter & ";samaccountname;subtree"
                    'MsgBox strQuery
                
                    'Com.ActiveConnection = Conn
                    ''Com.CommandText = strQuery
                    'Set Rs = Com.Execute
                    
                    'If Not Rs.EOF Then
                        'Rs.MoveFirst
                        
                        'While Not Rs.EOF
                        '    cbo.AddItem Rs.Fields(0)
                       ' Wend
                        'k = k + 1
                    'End If
                
                    'If k = 0 Then Exit Function
                End If
                
            Next
            
        End If
                            
        Set oTmpSG = Nothing
    Next
            
    Exit Function
    
errHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
    Resume Next
End Function




Sub HandleSingleField(indent, inx, f, strName)
     Dim t
     Dim bVerbose As Boolean
    
On Error GoTo Err:

     t = TypeName(f)
     
     If t = "String" Or t = "Boolean" Or t = "Date" Or t = "Long" Or t = "Byte" Then
     
        'If indent = 0 Then
            'List1.AddItem "Field#" & inx & " (" & strName & ", " & t & ") " & " " & f
        If indent > 0 Then
        
            'List1.AddItem Space(indent * 4) & "Index#" & inx & " (" & strName & ", " & t & ") " & " " & f
            bVerbose = True
            If strName = "homeMDBBL" And t = "String" And bVerbose Then
             ' f contains an adspath to a mailbox user
             DumpUserInfo f
            End If
            
        End If
        'List1.ListIndex = List1.ListCount - 1
        
        Exit Sub
     End If
    
    ' If Right (t, 2) = "()" Then
     If t = "Variant()" Then
        Dim i As Integer
    
        'List1.AddItem Space(indent * 4) & "Field#" & inx & " (" & strName & ", " & t & ") "
        
        For i = LBound(f) To UBound(f)
            Call HandleSingleField(indent + 1, i, f(i), strName)
        Next
        'List1.ListIndex = List1.ListCount - 1
        
        Exit Sub
     End If
    
     'List1.AddItem "Field#" & inx & " (" & strName & ", " & t & ") "
    Exit Sub
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
    Resume Next
End Sub

Function EnumerateFields(str, fieldlist)
     Dim inx, f, t
        
On Error GoTo Err:

     'List1.AddItem str & " count=" & fieldlist.Count
     'List1.ListIndex = List1.ListCount - 1
    
     inx = 0
     While inx < fieldlist.Count
      f = fieldlist(inx)
      'MsgBox fieldlist.Count
    
      Call HandleSingleField(0, inx, f, fieldlist(inx).name)
      inx = inx + 1
     Wend

    Exit Function
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
    Resume Next

End Function

Sub DumpUserInfo(strADSPath)
     Dim objPerson              'CDO.Person
     Dim objRecipient           'objPerson.GetInterface("IMailRecipient")  'As CDOEXM.IMailRecipient
     Dim sAlias As String
    
On Error GoTo Err:

    
     Set objPerson = CreateObject("CDO.Person")
     objPerson.DataSource.Open "LDAP://" & strADSPath
    
     Set objRecipient = objPerson.GetInterface("IMailRecipient")
    
    sAlias = objRecipient.Alias
    
     frmExMerge.List7.AddItem objRecipient.Alias & GetSize_Count(sCurrentServer, sAlias)
     
     
     'frmExMerge.List7.AddItem objRecipient
     'List1.AddItem vbTab & vbTab & "ForwardingStyle       = " & objRecipient.ForwardingStyle
     'List1.AddItem vbTab & vbTab & "ForwardTo             = " & objRecipient.ForwardTo
     'List1.AddItem vbTab & vbTab & "HideFromAddressBook   = " & objRecipient.HideFromAddressBook
     'List1.AddItem vbTab & vbTab & "IncomingLimit         = " & objRecipient.IncomingLimit
     'List1.AddItem vbTab & vbTab & "OutgoingLimit         = " & objRecipient.OutgoingLimit
     'List1.AddItem vbTab & vbTab & "ProxyAddresses        = " & sProxyAddresses(3, objRecipient.ProxyAddresses)
     'List1.AddItem vbTab & vbTab & "RestrictedAddresses   = " & objRecipient.RestrictedAddresses
     'List1.AddItem vbTab & vbTab & "RestrictedAddressList = " & sProxyAddresses(3, objRecipient.RestrictedAddressList)
     'List1.AddItem vbTab & vbTab & "SMTPEmail             = " & objRecipient.SMTPEmail
     'List1.AddItem vbTab & vbTab & "TargetAddress         = " & objRecipient.TargetAddress
    
     Set objRecipient = Nothing
     Set objPerson = Nothing

    Exit Sub
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
    Resume Next
End Sub
