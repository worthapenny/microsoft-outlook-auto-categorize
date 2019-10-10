Private WithEvents InboxItems As Outlook.Items
Private WithEvents SentItems As Outlook.Items

Private Sub Application_Startup()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    
    ' Initialize the list of known doamin dictionary
    AddDomainNames
    
    'Initialize the list of known subject lines
    AddSubjectLines
    
    ' Initialize the list of the know senders list
    AddSenders
    
    ' Create Categories based on the above dictionary
    CreateCategories objNS
    
    'default local Inbox
    Set InboxItems = objNS.GetDefaultFolder(olFolderInbox).Items
    Set SentItems = objNS.GetDefaultFolder(olFolderSentMail).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.mailItem
  If Item.Class = olMail _
        Or Item.Class = olMeetingRequest _
        Or Item.Class = olMeetingReceived Then
        
    OutlookItemReceivedOrSent Item
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  'MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub

Private Sub SentItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler
    
    If Item.Class = olMail _
        Or Item.Class = olMeetingRequest _
        Or Item.Class = olAppointmentItem Then
        
      OutlookItemReceivedOrSent Item
    End If
    
ProgramExit:
    Exit Sub
    
ErrorHandler:
    'MsgBox Err.Number & " - " & Err.Description
    Resume ProgramExit
End Sub


'' The Event handler below gets fired when a reminder
'' shows up in the Application.
'' It is used to send email to CSM
'' based on SCADAWatch Calendar
Private Sub Application_Reminder(ByVal Item As Object)
        
    '' Exit if it's not an Appointment
    If Item.MessageClass <> "IPM.Appointment" Then
        Exit Sub
    End If
    
    
    '' Exit if it's not an Appointment with RIGHT category
    If Item.Categories <> "AutoTrigger_SW_CalendarSummary" Then
        Exit Sub
    End If
    
    
    '' Send email out
    CalendarRelated.SendAutomatedEmailOutToCSMs Item
    
End Sub
