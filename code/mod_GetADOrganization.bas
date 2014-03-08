Attribute VB_Name = "mod_GetADOrganization"
Option Explicit

Function ADExists() As Boolean
    Dim iAdRootDSE As IADs
    
    ADExists = False
    
On Error GoTo Err:
    
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    ADExists = True
    Exit Function
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
End Function

Function Get_OrganizationName() As String
    Dim iAdRootDSE As ActiveDs.IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim varConfigNC As Variant
    Dim strQuery As String
    
    Get_OrganizationName = ""
    
On Error GoTo Err:
    
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    varConfigNC = iAdRootDSE.Get("configurationNamingContext")
    
    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"
    
    ' Build the query to find the organization.
    strQuery = "<LDAP://" & varConfigNC & ">;(objectCategory=msExchOrganizationContainer);name,cn,distinguishedName;subtree"
    
    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    While Not Rs.EOF
        Get_OrganizationName = Rs.Fields("cn")
        Rs.MoveNext
    Wend
    
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Com = Nothing
    Set Conn = Nothing
    Exit Function
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
End Function


Function Get_AdministrativeGroup() As String

    Dim iAdRootDSE As ActiveDs.IADs
    Dim Conn As New ADODB.Connection
    Dim Com As New ADODB.Command
    Dim Rs As ADODB.Recordset
    Dim varConfigNC As Variant
    Dim strQuery As String
    
On Error GoTo Err:
    
    ' Get the configuration naming context.
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    
    If iAdRootDSE Is Nothing Then
        Exit Function
    End If
    
    varConfigNC = iAdRootDSE.Get("configurationNamingContext")
    
    ' Open the connection.
    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"
    
    ' Build the query to find the organization.
    strQuery = "<LDAP://" & varConfigNC & ">;(objectCategory=msExchAdminGroup);name,cn,distinguishedName;subtree"
    
    Com.ActiveConnection = Conn
    Com.CommandText = strQuery
    Set Rs = Com.Execute
    
    ' Iterate through the results.
    While Not Rs.EOF
        
        ' Output the name of the organization.
        'MsgBox "The organization name is: " & Rs.Fields("cn")
        
        Get_AdministrativeGroup = Rs.Fields("cn")
        
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
    Debug.Print Err.Number & " - " & Err.Description
    If Err.Number = -2147023541 Then
        Resume Next
    End If
End Function
