Attribute VB_Name = "mod_GetADUsers"

Sub GetADUsers(DomainName As String)

    'DomainName is something like "DC=MYDOMAIN3,DC=example,DC=com"
    
    Dim objUser As IADsUser
    Dim objContainer As IADsContainer
    Dim objMailbox As CDOEXM.IMailboxStore
    Dim i As Long
    Dim name As String
    
    On Error GoTo Error
    ' get the container. Note that user information may be located in
    ' other organizational units.
    Set objContainer = GetObject("LDAP://CN=users," + DomainName)
    
    objContainer.Filter = Array("User")
    i = 0
    
    For Each objUser In objContainer
       name = objUser.name
       name = Right(name, Len(name) - 3)
       Set objMailbox = objUser
       If objMailbox.HomeMDB = "" Then
          'List1.AddItem name + "   (no mailbox)"
       Else
          'List1.AddItem name + "   (has mailbox)"
          'List1.AddItem objMailbox.HomeMDB
       End If
       i = i + 1
    Next
    'List1.AddItem "Number of users found in the DS (in the default container): " + str(i)
    GoTo Ending
    
Error:
       'List1.AddItem "Failed while displaying the users in the default container."
       MsgBox "Run time error: " + str(Err.Number) + " " + Err.Description
       Err.Clear
Ending:

End Sub



