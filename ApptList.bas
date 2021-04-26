Attribute VB_Name = "ApptList"
Sub CreateListofAppt()
   
   Dim CalFolder As Outlook.MAPIFolder
   Dim CalItems As Outlook.Items
   Dim ResItems As Outlook.Items
   Dim sFilter, strSubject, strAppt As String
   Dim iNumRestricted As Integer
   Dim itm, apptSnapshot As Object
   Dim tStart As Date, tEnd As Date, tFullWeek As Date
   Dim wd As Integer
  
   ' Use the default calendar folder
   Set CalFolder = Session.GetDefaultFolder(olFolderCalendar)
   Set CalItems = CalFolder.Items

   ' Sort all of the appointments based on the start time
   CalItems.Sort "[Start]"
   CalItems.IncludeRecurrences = True

   ' Set an end date
    tStart = Format(Date + 1, "Short Date")
    tEnd = Format(Date + 2, "Short Date")
    tFullWeek = Format(Date + 6, "Short Date")
 
    wd = Weekday(Date)
   ' Sun = 1, Mon = 2, Tues = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' get next day appt, do whole week on sunday
If wd >= 2 And wd <= 6 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tEnd & "'"
ElseIf wd = 1 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tFullWeek & "'"
End If

Debug.Print sFilter
   Set ResItems = CalItems.Restrict(sFilter)

   iNumRestricted = 0

   'Loop through the items in the collection.
   For Each itm In ResItems
   Debug.Print ResItems.Count
      iNumRestricted = iNumRestricted + 1
      
 ' Create list of appointments
  strAppt = strAppt & vbCrLf & itm.Subject & vbTab & " >> " & vbTab & itm.Start & vbTab & " to: " & vbTab & Format(itm.End, "h:mm AM/PM")

   Next
   
' After the last occurrence is checked
' Open a new email message form and insert the list of dates
  Set apptSnapshot = Application.CreateItem(olMailItem)
  With apptSnapshot
    .Body = strAppt & vbCrLf & "Total appointments; " & iNumRestricted
    .To = "me@slipstick.com"
    .Subject = "Appointments for " & tStart
    .Display 'or .send
  End With

   Set itm = Nothing
   Set apptSnapshot = Nothing
   Set ResItems = Nothing
   Set CalItems = Nothing
   Set CalFolder = Nothing
   
End Sub

Sub pruebita()

strStationeryFile = CStr(Environ("APPDATA")) & "\Microsoft\Stationery\Schnauzer.htm"
Debug.Print strStationeryFile

End Sub
