Sub MoveArchiveRootItemsToInbox()
    ' Macro to move all MailItems from the Online Archive root folder to its Inbox folder.
    ' This is useful for organizing archives where all items accumulate at the top level.

    Dim objOutlook As Object
    Dim objNamespace As Object
    Dim objArchiveStore As Object
    Dim objSourceRoot As Object
    Dim objDestinationInbox As Object
    Dim objItems As Object ' The collection of items in the source folder
    Dim objItem As Object
    Dim i As Long
    Dim lMovedCount As Long
    Dim lInitialCount As Long
    
    ' --- Configuration ---
    ' *** IMPORTANT: CHANGE THIS to your exact Online Archive display name ***
    Const ONLINE_ARCHIVE_NAME As String = "Online Archive - ghi.jkl@def.com" 
    
    ' Error handling is crucial in VBA, especially for server-side moves.
    On Error GoTo ErrorHandler

    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    lMovedCount = 0
    Debug.Print "--------------------------------------------------------"
    Debug.Print "Migration Start Time: " & Now()

    ' 1. Find the Online Archive Store
    Set objArchiveStore = FindStoreByName(objNamespace, ONLINE_ARCHIVE_NAME)
    
    If objArchiveStore Is Nothing Then
        ' Critical error, require user interaction
        MsgBox "CRITICAL ERROR: Online Archive '" & ONLINE_ARCHIVE_NAME & "' not found." & vbCrLf & _
               "Please ensure Outlook is open and the archive is connected.", vbCritical, "Migration Failed"
        GoTo Finalize
    End If
    
    ' 2. Define Source (Root) and Destination (Inbox)
    Set objSourceRoot = objArchiveStore.GetRootFolder ' The root folder of the archive
    
    ' Try to find the Inbox folder
    On Error Resume Next
    Set objDestinationInbox = objSourceRoot.Folders("Inbox")
    On Error GoTo ErrorHandler ' Reset error handler
    
    If objDestinationInbox Is Nothing Then
        ' Critical error, require user interaction
        MsgBox "CRITICAL ERROR: 'Inbox' folder not found under the archive root." & vbCrLf & _
               "Please ensure the Inbox exists in the Online Archive.", vbCritical, "Migration Failed"
        GoTo Finalize
    End If
    
    ' 3. Get items to move and confirm
    Set objItems = objSourceRoot.Items
    lInitialCount = objItems.Count
    
    If lInitialCount = 0 Then
        MsgBox "No items found in the root folder of the Online Archive. Migration complete.", vbInformation, "Migration Complete"
        GoTo Finalize
    End If
    
    Dim ConfirmationMsg As String
    ConfirmationMsg = "CONFIRMATION:" & vbCrLf & vbCrLf & _
                      "You are about to MOVE " & lInitialCount & " item(s) from the Archive Root" & vbCrLf & _
                      "to the Archive Inbox." & vbCrLf & vbCrLf & _
                      "Source: " & objSourceRoot.FolderPath & vbCrLf & _
                      "Destination: " & objDestinationInbox.FolderPath & vbCrLf & vbCrLf & _
                      "Do you want to proceed? This action cannot be undone."

    If MsgBox(ConfirmationMsg, vbYesNo + vbExclamation, "Confirm Move Operation") = vbNo Then
        MsgBox "Migration cancelled by user.", vbInformation
        GoTo Finalize
    End If
    
    ' 4. Perform the Move Operation (Iterate backwards)
    ' Iterating backwards is necessary when deleting/moving items from a collection.
    
    Debug.Print "Found " & lInitialCount & " items to move."
    Debug.Print "Starting move operation (check for ERROR logs below):"

    For i = lInitialCount To 1 Step -1
        ' --- RESILIENCE AGAINST THROTTLING/SERVER ERRORS ---
        ' Use On Error Resume Next during the move loop. 
        ' This allows server throttling errors to be logged without stopping the entire script.
        On Error Resume Next 
        
        Set objItem = objItems.Item(i)
        
        ' Only attempt to move MailItems (Mails, Meeting Requests, etc., Class 43)
        If objItem.Class = olMail Then
            objItem.Move objDestinationInbox
            
            If Err.Number <> 0 Then
                ' Log the error but don't stop the script (resilience against throttling)
                Debug.Print "ERROR moving item '" & objItem.Subject & "'. Error: " & Err.Description & " (Item " & i & "/" & lInitialCount & ")"
                Err.Clear ' Clear the error state to allow the loop to continue
            Else
                lMovedCount = lMovedCount + 1
                ' Log successful moves periodically to show progress
                If lMovedCount Mod 100 = 0 Then
                    Debug.Print "Progress: Successfully moved " & lMovedCount & " items so far."
                End If
            End If
        Else
            ' Log non-mail items that were skipped
            Debug.Print "Skipping non-MailItem (Class: " & objItem.Class & "): " & objItem.Subject
        End If
        
        On Error GoTo ErrorHandler ' Reset to general error handler for final checks
        
    Next i
    
    ' 5. Final Report
    Debug.Print "Migration End Time: " & Now()
    Debug.Print "--------------------------------------------------------"
    
    MsgBox "Migration complete!" & vbCrLf & _
           "Total items detected: " & lInitialCount & vbCrLf & _
           "Successfully moved: " & lMovedCount & vbCrLf & _
           "Failed/Skipped: " & (lInitialCount - lMovedCount) & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G in VBA Editor) for detailed error logs.", vbInformation, "Migration Summary"

Finalize:
    ' Clean up objects and status bar
    Set objItems = Nothing
    Set objSourceRoot = Nothing
    Set objDestinationInbox = Nothing
    Set objArchiveStore = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    Exit Sub

ErrorHandler:
    ' This handles errors outside of the resilient item moving loop 
    MsgBox "An unrecoverable error occurred: " & Err.Description & " (Number: " & Err.Number & ")", vbCritical, "Operation Error"
    Resume Finalize

End Sub

Function FindStoreByName(ByVal objNamespace As Object, ByVal strStoreName As String) As Object
    ' Helper function to find a store (mailbox/PST) by its DisplayName
    Dim objStore As Object
    
    For Each objStore In objNamespace.Stores
        If objStore.DisplayName = strStoreName Then
            Set FindStoreByName = objStore
            Exit Function
        End If
    Next objStore
    
    Set FindStoreByName = Nothing
End Function
