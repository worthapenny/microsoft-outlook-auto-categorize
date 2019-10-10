Public KnownDomains As Dictionary
Public KnownSubjectLines As Dictionary
Public KnownSenders As Dictionary


Public Sub AddSenders()
    Set KnownSenders = New Dictionary
    KnownSenders.Add "Help@domain.com", "Help"
    KnownSenders.Add "do-not-reply@do.not-reply.com", "NDR"
End Sub


Public Sub AddSubjectLines()
    Set KnownSubjectLines = New Dictionary
    KnownSubjectLines.Add "Subject line", "SL"
    
End Sub
   

Public Sub AddDomainNames()
    Set KnownDomains = New Dictionary
    KnownDomains.Add "twitter.com", "Tweet"
    KnownDomains.Add "gmail.com", "Gmail"
    
End Sub

Public Sub CreateCategories(mapiNameSpace As Outlook.NameSpace)
    Dim oCategory As Category
    Dim name As String
    Dim domainItem, subjectItem, senderItem As Variant
    Dim categoryExists As Boolean
    
    
    ' Add categories for known domains
    For Each domainItem In KnownDomains.KeyValuePairs
        AddCategory mapiNameSpace, domainItem.value
    Next
    
    
    ' Add categories for known domains
    For Each subjectItem In KnownSubjectLines.KeyValuePairs
        AddCategory mapiNameSpace, subjectItem.value
    Next
    
    
    ' Add categories for known senders
    For Each senderItem In KnownSenders.KeyValuePairs
        AddCategory mapiNameSpace, senderItem.value
    Next
    
End Sub

Public Sub AddCategory(mapiNameSpace As Outlook.NameSpace, categoryName As String)
    Dim categoryExists As Boolean
    categoryExists = False
    
    ' check if category already exists
    For Each oCategory In mapiNameSpace.Categories
       categoryExists = categoryExists Or (oCategory.name = categoryName)
    Next
    
     ' if category does not exists create one
    If Not categoryExists Then
        mapiNameSpace.Categories.Add (categoryName)
    End If
    
    
End Sub

Public Sub OutlookItemReceivedOrSent(mailItem As Object)
    Dim emailAddresses As Collection: Set emailAddresses = GetEmails(mailItem)
    ApplyCategory emailAddresses, mailItem
    
    
    Dim emailSubject As String: emailSubject = mailItem.Subject
    ApplyCategoryBySubject emailSubject, mailItem
    
    
    Dim sender As String: sender = GetSenderEmail(mailItem)
    ApplyCategoryBySender sender, mailItem
    
End Sub


Public Sub ApplyCategoryBySender(emailSender As String, mailItem As Object)
    Dim senderItem As Variant
    For Each senderItem In KnownSenders.KeyValuePairs
        If senderItem.key = emailSender Then
            ApplyCategoryToEmail mailItem, senderItem.value
        End If
        
    Next
    
End Sub


Public Sub ApplyCategoryBySubject(emailSubject As String, mailItem As Object)
    Dim knownSubject As Variant
    For Each knownSubject In KnownSubjectLines.KeyValuePairs
        
        If InStr(emailSubject, knownSubject.key) > 0 Then
            ApplyCategoryToEmail mailItem, knownSubject.value
        End If
        
    Next
    
End Sub

Public Sub ApplyCategory(emailAddresses As Collection, emailItem As Object)
    Dim email As Variant
    For Each email In emailAddresses
        Dim emailAddress As String: emailAddress = CStr(email)
        If ValidEmail(emailAddress) Then
            
            Dim domainName As String
            domainName = GetDomain(emailAddress)
            
            Dim domainItems As Variant
            For Each domainItems In KnownDomains.KeyValuePairs
                If domainName = domainItems.key Then
                    ApplyCategoryToEmail emailItem, domainItems.value ' value = category name
                End If
            Next
            
        End If
    Next
End Sub

Private Sub ApplyCategoryToEmail(mailItem As Object, categoryName As String)
    Dim Exists As Boolean
    Dim arrCategories As Variant

    ' Initialize variables
    Exists = False
    arrCategories = Split(mailItem.Categories, ", ")

    ' Loop through all categories
    For i = LBound(arrCategories) To UBound(arrCategories)

        ' Check if the specified category already exists
        If StrComp(categoryName, arrCategories(i)) = 0 Then
            Exists = True
            Exit For
        End If
    Next i

    ' If the category does not exist, add it
    If Not Exists Then
        If Len(mailItem.Categories) > 0 Then
            mailItem.Categories = mailItem.Categories & ", " & categoryName
        Else
             mailItem.Categories = categoryName
        End If
        mailItem.Save
        
        'MsgBox "Category applied: " & categoryName
    End If
End Sub

Public Function ValidEmail(emailAddress As String) As Boolean
    Dim oRegEx As Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx
        .Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
        ValidEmail = .test(emailAddress)
    End With
    Set oRegEx = Nothing
End Function

Public Function GetDomain(emailAddress As String) As String
    Dim brokenEmail() As String: brokenEmail = Split(emailAddress, "@")
    If UBound(brokenEmail) = 1 Then
        GetDomain = brokenEmail(1)
    End If
End Function

Public Function GetSenderEmail(email As Object)
    Dim senderEmail As String
    
    ' work with the sender
    On Error GoTo EndOfMethod
    If email.SenderEmailType = "EX" Then
        
        If email.Class = olMeetingRequest Then
            ' there is no email.Recipient for this this class
            ' hence get the recipent by replying that item
            Dim replyEmail As mailItem: Set replyEmail = email.Reply()
            Dim contact As Recipient: Set contact = replyEmail.Recipients.Item(1)
            replyEmail.Close olDiscard
            
            senderEmail = getExchangeEmailAddress(contact)
        Else
            senderEmail = getExchangeEmailAddress(email.sender)
        End If
    Else
        senderEmail = email.SenderEmailAddress
    End If
    
EndOfMethod:
    
    GetSenderEmail = senderEmail
End Function

Public Function GetEmails(email As Object) As Collection
    Set GetEmails = New Collection
    
    'Get Sender's email
    GetEmails.Add (GetSenderEmail(email))
    

    On Error GoTo SkipAllRecipients
    ' work with the recipients
    Dim receiver As Recipient
    
    For Each receiver In email.Recipients
        On Error GoTo SkipRecipient
        If InStr(receiver.Address, "Exchange") Then
            GetEmails.Add (getExchangeEmailAddress(receiver))
        Else
            GetEmails.Add receiver.Address
        End If

SkipRecipient:
    Next
    
SkipAllRecipients:
    
End Function

Private Function getExchangeEmailAddress(contact As Object) As String
    On Error GoTo err
    Dim possibleEmail As String: possibleEmail = contact.Address
    If ValidEmail(possibleEmail) Then
        getExchangeEmailAddress = possibleEmail
    Else
        Dim PR_SMTP_ADDRESS As String: PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        getExchangeEmailAddress = contact.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    End If
err:
End Function

'' ****************************************
'' - - - - - - - - - - - - - - - - - - - -
''
'' TEST SUBS
''
'' - - - - - - - - - - - - - - - - - - - -
'' ****************************************

Public Sub ApplyCategoriesToAllInboxItems()
    Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
    Dim Item As Object
    
    ' Initialize the list of known doamin dictionary
    AddDomainNames
    
    ' Initialize the list of known subject lines
    AddSubjectLines
    
    ' Initialize the list of the know senders list
    AddSenders
    
    For Each Item In olFolder.Items
        If Item.Class = olMail _
        Or Item.Class = olMeetingReceived Then
        
            OutlookItemReceivedOrSent Item
        End If
    Next
    
    MsgBox "Done"
End Sub


Public Sub ApplyCategoriesToAllSentItems()
    Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    Set olFolder = objNS.GetDefaultFolder(olFolderSentMail)
    Dim Item As Object
    
    ' Initialize the list of known doamin dictionary
    AddDomainNames
    
    ' Initialize the list of known subject lines
    AddSubjectLines
    
    ' Initialize the list of the know senders list
    AddSenders
    
    For Each Item In olFolder.Items
        If Item.Class = olMail _
        Or Item.Class = olMeetingReceived Then
        
            OutlookItemReceivedOrSent Item
        End If
    Next
    
    MsgBox "Done"
End Sub


Private Sub test()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    
    ' Initialize variables
    AddDomainNames
    
    ' Initialize the list of known subject lines
    AddSubjectLines
    
    ' Initialize the list of the know senders list
    AddSenders
    
    Dim testFolder As Outlook.Folder: Set testFolder = objNS.GetDefaultFolder(olFolderInbox).Folders("test")
    'Dim testFolder As Outlook.Folder: Set testFolder = objNS.GetDefaultFolder(olFolderInbox)
    Dim Item As Object
    
    For Each Item In testFolder.Items
        OutlookItemReceivedOrSent Item
    Next
    
End Sub
