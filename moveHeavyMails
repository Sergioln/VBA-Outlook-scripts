Sub moveHeavyMails()

    'Create the main object to manage theOutlook session
    Dim otlkOject As Outlook.NameSpace: Set otlkObject = GetNamespace("MAPI")

    Dim srcFolder As Outlook.MAPIFolder
    Dim dstFolder As Outlook.MAPIFolder

    Dim item As Object
    Dim nMails As Integer: nMails = 0
    
    'Set the source Mailbox or PST name
        'srcMailBoxName = "Origen Prueba"
        'src_Pst_Folder_Name = "Bandeja de entrada"
        'Set srcFolder = Outlook.Session.Folders(srcMailBoxName).Folders(src_Pst_Folder_Name)
    'Get the default inbox folder
    Set srcFolder = otlkObject.GetDefaultFolder(olFolderInbox)
    
    'Set the destination Mailbox or PST name
    dstMailBoxName = "Archivo 1"
    dst_Pst_Folder_Name = "Grandes"
    Set dstFolder = Outlook.Session.Folders(dstMailBoxName).Folders(dst_Pst_Folder_Name)
    
    'Loop through all source folfer items
    For Each item In srcFolder.Items
        If TypeOf item Is Outlook.MailItem Then
            Dim currentMail As Outlook.MailItem: Set currentMail = item
            'Date of receving
            MsgBox (currentMail.ReceivedTime)
            'Check it´s older than yesterday (-1) and bigger than 2 MB and it´s read
            If (currentMail.ReceivedTime < (DateTime.now - 1)) And (currentMail.Size > 2000000) And (Not currentMail.UnRead) Then
                
                'Size of the total message
                MsgBox "Total size: " & currentMail.Size & "\nItem size: " & currentMail.Attachments.item(1).Size
                nMails = nMails + 1
                currentMail.Move dstFolder
            End If
        'If n > 5 Then
            'Exit For
        End If
    Next
    MsgBox SourceFolder & nMails & " correos pesados movidos"
End Sub
