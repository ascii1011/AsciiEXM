Attribute VB_Name = "mod_GetDomainGroupsUsers"


Function GetUserDetails(sUser As String, i As Integer)
    Dim iAdRootDSE As ActiveDs.IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim varConfigNC As Variant
    Dim strQuery As String
    
    On Error GoTo Err:
    
    ' Get the configuration naming context.
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    varConfigNC = iAdRootDSE.Get("defaultNamingContext")
    'varConfigNC = iAdRootDSE.Get("configurationNamingContext")
    
    ' Open the connection.
    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"
    
    ' Build the query to find the organization.
    'strQuery = "<LDAP://" & varConfigNC & ">;(uid=" & sUser & ");cn,name,samaccountname,mail,telephoneNumber;subtree"
    strQuery = "<LDAP://" & varConfigNC & ">;(mailNickName=" & sUser & ");adspath,cn,name,samaccountname,mail,telephoneNumber;subtree"
    
    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    ' Iterate through the results.
    While Not Rs.EOF
        
        ' Output the name of the organization.
        'MsgBox "The organization name is: " & Rs.Fields("cn")
        Main.ADUsers(i).CN = Rs.Fields("cn").value
        Main.ADUsers(i).SamAccountName = Rs.Fields("samaccountname").value
        
        'List1.AddItem Rs.Fields("mail").Value
        'List1.AddItem Rs.Fields("telephoneNumber").Value
        
        Rs.MoveNext
        
    Wend
    
    'Clean up.
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Com = Nothing
    Set Conn = Nothing
    Exit Function
    
Err:
    MsgBox Err.Number & vbNewLine & Err.Description
    
    Exit Function

End Function

Sub UserInfo3()

    sUser = "charty"
    sDN = "cn=" & sUser & ",ou=people,dc=siboneyusa,dc=com"
    sRoot = "LDAP://susany0.siboneyusa.com/dc=siboneyusa,dc=com"
    
    Dim oDS: Set oDS = GetObject("LDAP:")
    Dim oAuth: Set oAuth = oDS.OpenDSObject(sRoot, , sDN, "password", &H200)
    
    Dim oConn: Set oConn = CreateObject("ADODB.Connection")
    oConn.Provider = "ADSDSOObject"
    oConn.Open "Ads Provider", sDN, "password"
    
    Dim Rs
    Set Rs = oConn.Execute("<" & sRoot & ">;(uid=" & sUser & ");cn,mail,telephoneNumber;subtree")
    
    'List1.AddItem Rs("cn").value
    'List1.AddItem Rs("mail").value
    'List1.AddItem Rs("telephoneNumber").value
End Sub


Sub UserInfo2()

    sUser = "charty"
    sDN = "cn=" & sUser & ",ou=people,dc=company,dc=com"
    sRoot = "LDAP://ldapservername.com/dc=company,dc=com"
    
    Dim oDS: Set oDS = GetObject("LDAP:")
    Dim oAuth: Set oAuth = oDS.OpenDSObject(sRoot, sDN, "password", &H200)
    
    Dim oConn: Set oConn = CreateObject("ADODB.Connection")
    oConn.Provider = "ADSDSOObject"
    oConn.Open "Ads Provider", sDN, "password"
    
    Dim Rs
    Set Rs = oConn.Execute("<" & sRoot & ">;(uid=" & sUser & ");cn,mail,telephoneNumber;subtree")
    
    'List1.AddItem Rs("cn").value
    'List1.AddItem Rs("mail").value
    'List1.AddItem Rs("telephoneNumber").value
End Sub



Sub ListDomains()
    Dim objNameSpace As Object
    Dim Domain, objDomain
    Dim strDomains() As String
    Dim sTmp As String
    Dim i As Integer
    
    Set objNameSpace = GetObject("WinNT:")
    
    i = 0
    For Each objDomain In objNameSpace
        sTmp = sTmp & objDomain.name & ","
        i = i + 1
    Next
    
    strDomains = Split(sTmp, ",")
    
    ReDim Main.Domains(UBound(strDomains))
    If UBound(strDomains) > 0 Then
    
        For i = 0 To UBound(strDomains) - 1
            Main.Domains(i).name = strDomains(i)
            ListUsers Main.Domains(i)
            
            'ListGroups objDomain.Name
        Next i
        
    End If
    
End Sub

'Sub ListGroups(strDomain)
'    Set objComputer = GetObject("WinNT://" & strDomain)
'    objComputer.Filter = Array("Group")
'    For Each objgroup In objComputer
'        List1.AddItem vbTab & "Group: " & objgroup.Name
'        Debug.Print objgroup.Name
'    Next
'End Sub

Sub LinkLDAPAccounts()
    Dim i As Integer
    Dim j As Integer
    
    
    For i = 0 To UBound(Main.ADUsers) - 1
    
        'for j = 0 to ubound(main.Exch.Svrs(main.Current.
    
    
    
    
    Next i
    
    
End Sub


Sub ListUsers(oDomain As Domains_Struct)
    'Dim objComputer As IADsMembers
    Dim objComputer
    Dim objUser As IADsUser
    Dim i As Integer
    
    Set objComputer = GetObject("WinNT://" & oDomain.name)
    objComputer.Filter = Array("User")
    
    ReDim oDomain.ADUsers(objComputer.Count)
    i = 0
    For Each objUser In objComputer
        oDomain.ADUsers(i).Domain = strDomain
        oDomain.ADUsers(i).AccountDisabled = objUser.AccountDisabled
        oDomain.ADUsers(i).AccountExpirationDate = objUser.AccountExpirationDate
        oDomain.ADUsers(i).ADsPath = objUser.ADsPath
        oDomain.ADUsers(i).BadLoginAddress = objUser.BadLoginAddress
        oDomain.ADUsers(i).BadLoginCount = objUser.BadLoginCount
        oDomain.ADUsers(i).Class = objUser.Class
        oDomain.ADUsers(i).Department = objUser.Department
        oDomain.ADUsers(i).Description = objUser.Description
        oDomain.ADUsers(i).Division = objUser.Division
        oDomain.ADUsers(i).EmailAddress = objUser.EmailAddress
        oDomain.ADUsers(i).EmployeeID = objUser.EmployeeID
        oDomain.ADUsers(i).FaxNumber = objUser.FaxNumber
        oDomain.ADUsers(i).FirstName = objUser.FirstName
        oDomain.ADUsers(i).FullName = objUser.FullName
        oDomain.ADUsers(i).GraceLoginsAllowed = objUser.GraceLoginsAllowed
        oDomain.ADUsers(i).GraceLoginsRemaining = objUser.GraceLoginsRemaining
        oDomain.ADUsers(i).Guid = objUser.Guid
        oDomain.ADUsers(i).HomeDirectory = objUser.HomeDirectory
        oDomain.ADUsers(i).HomePage = objUser.HomePage
        oDomain.ADUsers(i).IsAccountLocked = objUser.IsAccountLocked
        oDomain.ADUsers(i).Languages = objUser.Languages
        oDomain.ADUsers(i).LastFailedLogin = objUser.LastFailedLogin
        oDomain.ADUsers(i).LastLogin = objUser.LastLogin
        oDomain.ADUsers(i).LastLogoff = objUser.LastLogoff
        oDomain.ADUsers(i).LastName = objUser.LastName
        oDomain.ADUsers(i).LoginHours = objUser.LoginHours
        oDomain.ADUsers(i).LoginScript = objUser.LoginScript
        oDomain.ADUsers(i).LoginWorkstations = objUser.LoginWorkstations
        oDomain.ADUsers(i).Manager = objUser.Manager
        oDomain.ADUsers(i).MaxLogins = objUser.MaxLogins
        oDomain.ADUsers(i).MaxStorage = objUser.MaxStorage
        oDomain.ADUsers(i).name = objUser.name
        oDomain.ADUsers(i).NamePrefix = objUser.NamePrefix
        oDomain.ADUsers(i).NameSuffix = objUser.NameSuffix
        oDomain.ADUsers(i).OfficeLocations = objUser.OfficeLocations
        oDomain.ADUsers(i).OtherName = objUser.OtherName
        oDomain.ADUsers(i).Parent = objUser.Parent
        oDomain.ADUsers(i).PasswordExpirationDate = objUser.PasswordExpirationDate
        oDomain.ADUsers(i).PasswordLastChanged = objUser.PasswordLastChanged
        oDomain.ADUsers(i).PasswordMinimumLength = objUser.PasswordMinimumLength
        oDomain.ADUsers(i).PasswordRequired = objUser.PasswordRequired
        oDomain.ADUsers(i).Picture = objUser.Picture
        oDomain.ADUsers(i).PostalAddresses = objUser.PostalAddresses
        oDomain.ADUsers(i).PostalCodes = objUser.PostalCodes
        oDomain.ADUsers(i).Profile = objUser.Profile
        oDomain.ADUsers(i).RequireUniquePassword = objUser.RequireUniquePassword
        oDomain.ADUsers(i).Schema = objUser.Schema
        oDomain.ADUsers(i).SeeAlso = objUser.SeeAlso
        oDomain.ADUsers(i).TelephoneHome = objUser.TelephoneHome
        oDomain.ADUsers(i).TelephoneMobile = objUser.TelephoneMobile
        oDomain.ADUsers(i).TelephoneNumber = objUser.TelephoneNumber
        oDomain.ADUsers(i).TelephonePager = objUser.TelephonePager
        oDomain.ADUsers(i).Title = objUser.Title
        
        GetUserDetails objUser.name, i
    Next
End Sub













