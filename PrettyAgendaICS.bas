Attribute VB_Name = "PrettyAgendaICS"
Public Sub SendPrettyAgenda()
Dim oNamespace As NameSpace
Dim oFolder As Folder
Dim oCalendarSharing As CalendarSharing
Dim objMail As MailItem
Dim wd As Integer

Set oNamespace = Application.GetNamespace("MAPI")
Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar)
Set oCalendarSharing = oFolder.GetCalendarExporter

' get the day - send sat/sun/monday out Fri night
' Sun = 1, Mon = 2, Tue = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' none set Sat/Sun

'wd = Weekday(Date)
'If wd >= 2 And wd <= 6 Then
'    lDays = Date + 1
'ElseIf wd = 1 Then
'    lDays = Date + 7
'End If

With oCalendarSharing
' options are olFreeBusyAndSubject, olFullDetails, olFreeBusyOnly
    .CalendarDetail = olFreeBusyAndSubject
    .IncludeWholeCalendar = False
    .IncludeAttachments = False
    .IncludePrivateDetails = True
    .RestrictToWorkingHours = False
    .StartDate = Date '  + 1
    .EndDate = Date 'lDays in case wd section is used
End With

' prepare as email
' options: olCalendarMailFormatEventList, olCalendarMailFormatDailySchedule
Set objMail = oCalendarSharing.ForwardAsICal(olCalendarMailFormatDailySchedule)
 
 ' Send the mail item to the specified recipient.
 With objMail
 .Recipients.Add "felix.reta@gmail.com"
 .Subject = "Updated Office calendar"

' Remove the attached ics if necessary
' .Attachments.Remove (1)
 .Display 'for testing, change to .send
 End With

Set oCalendarSharing = Nothing
Set oFolder = Nothing
Set oNamespace = Nothing
End Sub

