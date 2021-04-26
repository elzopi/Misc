Attribute VB_Name = "AccFwdAppt"
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
    
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
        ' Selected
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
        ' Open
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    ' Identify Message class ID
    Debug.Print GetCurrentItem.MessageClass
    Debug.Print Environ("temp")
    
    If GetCurrentItem.MessageClass = "IPM.Schedule.Meeting.Request" Then
       GetCurrentItem.SaveAs Environ("temp") & "\" & "TCal", olHTML
       
    End If
    Set objApp = Nothing
End Function

Sub ChangeMeeting()

Dim oRequest As MeetingItem
Dim oAppt As AppointmentItem

Set oRequest = Application.ActiveExplorer.Selection.Item(1)
If oRequest.MessageClass = "IPM.Schedule.Meeting.Request" Then
   Set oAppt = oRequest.GetAssociatedAppointment(True)
  
' set fields on the appt.
With oAppt
       .ReminderMinutesBeforeStart = 1080
       .Categories = "Slipstick"
       .ReminderSet = True
       .BusyStatus = olOutOfOffice
       .Save ' use .Display if you want to see the appt. and set the reminder yourself
  End With
  
End If

' use this to autoaccept
Dim oResponse
 Set oResponse = oAppt.Respond(olMeetingAccepted, True)
 oResponse.sEnd

'delete the request from the inbox
oRequest.Delete

End Sub


Sub AcceptAndForward()
 
Dim oAppt As AppointmentItem
Dim cAppt As AppointmentItem
Dim meAttendee As Outlook.Recipient
Dim oResponse
 
Set cAppt = GetCurrentItem.GetAssociatedAppointment(True)
Set oAppt = Application.CreateItem(olAppointmentItem)
 
With oAppt
    .MeetingStatus = olMeeting
    .Subject = "Accepted: " & cAppt.Subject
    .Start = cAppt.Start
    .Duration = cAppt.Duration
    .Location = cAppt.Location
    Set meAttendee = .Recipients.Add("me@mydomain.com")
     meAttendee.Type = olRequired
    .sEnd
End With
 
Set oResponse = cAppt.Respond(olMeetingAccepted, True)
oResponse.sEnd
 
Set cAppt = Nothing
Set oAppt = Nothing
 
End Sub
