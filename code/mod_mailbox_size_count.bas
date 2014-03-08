Attribute VB_Name = "mod_mailbox_size_count"
Function GetSize_Count(ServerName As String, MailBox As String) As String
   Dim oSession
   Dim oInfoStores
   Dim oInfoStore
   Dim StorageUsed
   Dim NumMessages
   Dim strProfileInfo
   Dim sMsg
   Const OneKiloByte = 1024


On Error GoTo Err:

   'Create Session object.
   Set oSession = CreateObject("MAPI.Session")
   If Err.Number <> 0 Then
      sMsg = "Error creating MAPI.Session."
      sMsg = sMsg & "Make sure CDO 1.21 is installed. "
      sMsg = sMsg & Err.Number & " " & Err.Description
      Debug.Print sMsg
      Exit Function
   End If
    
   strProfileInfo = ServerName & vbLf & MailBox

   'Log on.
   oSession.Logon , , False, True, , True, strProfileInfo
   If Err.Number <> 0 Then
      sMsg = "Error logging on: "
      sMsg = sMsg & Err.Number & " " & Err.Description
      Debug.Print sMsg
      Debug.Print "Server: " & ServerName
      Debug.Print "Mailbox: " & MailBox
      Set oSession = Nothing
      Exit Function
   End If

   'Grab the information stores.
   Set oInfoStores = oSession.InfoStores
   If Err.Number <> 0 Then

      sMsg = "Error retrieving InfoStores Collection: "
      sMsg = sMsg & Err.Number & " " & Err.Description
      Debug.Print sMsg
      Debug.Print "Server: " & ServerName
      Debug.Print "Mailbox: " & MailBox
      Set oInfoStores = Nothing
      Set oSession = Nothing
      Exit Function
   End If
    
   'Loop through information stores to find the user's mailbox.
   For Each oInfoStore In oInfoStores
      If InStr(1, oInfoStore.name, "Mailbox - ", 1) <> 0 Then
         '&HE080003 = PR_MESSAGE_SIZE '&HE080003
         StorageUsed = CLng(oInfoStore.Fields(&HE080014)) / OneKiloByte
         If Err.Number <> 0 Then
            sMsg = "Error retrieving PR_MESSAGE_SIZE: "
            sMsg = sMsg & Err.Number & " " & Err.Description
            Debug.Print sMsg
            Debug.Print "Server: " & ServerName
            Debug.Print "Mailbox: " & MailBox
            Set oInfoStore = Nothing
            Set oInfoStores = Nothing
            Set oSession = Nothing
            Exit Function
         End If
         
         '&H33020003 = PR_CONTENT_COUNT
         NumMessages = oInfoStore.Fields(&H36020003)

         If Err.Number <> 0 Then

            sMsg = "Error Retrieving PR_CONTENT_COUNT: "
            sMsg = sMsg & Err.Number & " " & Err.Description
            Debug.Print sMsg
            Debug.Print "Server: " & ServerName
            Debug.Print "Mailbox: " & MailBox
            Set oInfoStore = Nothing
            Set oInfoStores = Nothing
            Set oSession = Nothing
            Exit Function
         End If

        GetSize_Count = vbTab & "| " & NumMessages & "| " & StorageUsed
         sMsg = "Storage Used in " & oInfoStore.name
         sMsg = sMsg & " (bytes): " & StorageUsed
         Debug.Print sMsg
         Debug.Print "Number of Messages: " & NumMessages
      End If
   Next

   ' Log off.
   oSession.Logoff

   ' Clean up memory.
   Set oInfoStore = Nothing
   Set oInfoStores = Nothing
   Set oSession = Nothing
    Exit Function
    
Err:
    Debug.Print Err.Number & " - " & Err.Description
    Resume Next
End Function
