VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "ThisOutlookSession"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True
Public WithEvents myitem As MailItem
Attribute myitem.VB_VarHelpID = -1
'Attribute myitem.VB_VarHelpID = -1
Public WithEvents myMail As MailItem
Attribute myMail.VB_VarHelpID = -1
'Attribute myMail.VB_VarHelpID = -1
Public gcolMyInspectors As Collection
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1
'Attribute Items.VB_VarHelpID = -1

Public Sub Initialize_Handler()
 
 Set myMail = Application.ActiveInspector.CurrentItem
 Set myOlApp = Outlook.Application 'manipulate sent items NOT WORKING TO DATE
 
End Sub
Private Sub Application_Startup()
   Dim Ns As Outlook.NameSpace
   Set Ns = Application.GetNamespace("MAPI")
   Set curCal = Ns.GetDefaultFolder(olFolderCalendar).Items
   Set cureMail = Ns.GetDefaultFolder(olFolderInbox).Items
   Set Ns = Nothing
   inProgress = False
   Set objReminders = Application.Reminders
   
' Added to catch when mail is flagged up as Follow Up
   Set gcolMyInspectors = New Collection
   Set colInspectors = Application.Inspectors

End Sub
Private Sub colInspectors_NewInspector(ByVal Inspector As Inspector)

 ' This will be called everytime we open
 ' a new Inspector, so check if this is
 ' one that we want to monitor
 
 Dim MyInspectorHandler As clsInspectorHandler

 Es = Inspector.CurrentItem.Class

    msgbox ("At NewInspector Class: " & Es)

' If Inspector.currentItem.Class = olMail Then
'     If Not Inspector.currentItem.Sent Then
'         ' This is an unsent email so we want to
'         ' trap the buttons that can be clicked
'         Set MyInspectorHandler = New clsInspectorHandler
'         Call MyInspectorHandler.SetInspector(Inspector)
'         gcolMyInspectors.Add MyInspectorHandler
'     End If
' End If

End Sub

Private Sub myMail_Reply(ByVal Response As Object, Cancel As Boolean)

    msgbox "this is a reply"

' Set Response.SaveSentMessageFolder = myItem.Parent
 
End Sub
'Categorize Sent Items
'Place in ThisOutlookSession
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

' Code developed to capture sent messages and assign a Category, omits automated messages from Slack
' Adding functionality to capture Meeting related messages to:
'   a) Save accepted meeting requests
'   b) Send copy of the meeting to my personal email to have mobile consistent calendar
'   c) WIP avoid sending declined invites IPM........


Dim Recips As Recipients
Dim strDomain As String
Dim bOne As Boolean
Dim objApp As Application
Dim getCurrItem As Object
Dim oMail As MailItem
Dim mtgItem As MeetingItem
Dim oAppt As AppointmentItem
Dim strApptFile As String
Dim strStationeryFile As String
Dim objFileSystem As Object
Dim objTextStream As Object
Dim strHTMLBody As String
Dim strTextStream As String

    Set objApp = Application
    bOne = True
' Ignore if forwarded to slack.com channel, e.g. has magicleap.slack.com as part of the To email address oMail.Recipients

    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
        ' Selected
            Set getCurrItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
        ' Open
            Set getCurrItem = objApp.ActiveInspector.CurrentItem
    End Select
    ' Identify Message class ID
    Debug.Print getCurrItem.MessageClass

If TypeOf getCurrItem Is Outlook.MeetingItem And bOne Then
    
    Debug.Print "Meeting Item being processed"
 '   Set getCurrItem = objApp.acitveinspector.CurrentItem
'    Set getCurrItem = Item.Subject
    strApptFile = Environ("temp") & "\" & "TCal"
    Set mtgItem = getCurrItem
    mtgItem.ShowCategoriesDialog
    bOne = False
    Debug.Print mtgItem.MessageClass
    Debug.Print "File is: " & strApptFile
    mtgItem.SaveAs strApptFile & ".html", olHTML
    Set oAppt = mtgItem.GetAssociatedAppointment(True)
    oAppt.SaveAs strApptFile & ".ics", olICal
    
' WIP trying to get stationery into the mail body with appointment information
' Since we are dealing with HTML code, stationery has to be part of the appointment HTML

'Change the path to the specific stationery file
'             strStationeryFile = CStr(Environ("USERPROFILE")) & "\AppData\Roaming\Microsoft\StationeryShelly.htm"
'    strStationeryFile = CStr(Environ("APPDATA")) & "\Microsoft\Stationery\Schnauzer.htm"
'    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'    Set objTextStream = objFileSystem.OpenTextFile(strStationeryFile)
'    strTextStream = objTextStream.ReadAll

' Create email using Stationery plus invite msg body and attaching .ics above
    Set oMail = objApp.CreateItem(olMailItem)
    With oMail
        
        .To = "felix.reta@gmail.com"
        .Subject = mtgItem.Subject
        .Importance = olImportanceNormal
        .Attachments.Add strApptFile & ".ics"
        .BodyFormat = olFormatHTML
'        .HTMLBody = strTextStream ' testing to insert stationery
'        .HTMLBody = mtgItem.Body
        .Body = mtgItem.Body

'        .Body = mtgItem.Body
        .sEnd ' or
'        .Display ' for testing
    
    End With

Else

    If TypeOf Item Is Outlook.MailItem And Len(Item.Categories) = 0 Then
       Set Recips = Item.Recipients
'       Debug.Print Recips.Count
       bOne = True
       
       For Each Recipient In Item.Recipients
           Debug.Print Recipient.address
           strDomain = stringAfter(Recipient.address, "@")
'           Debug.Print strDomain
       
           If strDomain = "magicleap.slack.com" Then
              Item.Categories = "Maintenance"
           Else
                If bOne Then
           ' Set Item = Application.ActiveInspector.currentItem ' to open the Categories Dialog
                    Item.Categories = ""
                    Item.ShowCategoriesDialog
                    bOne = False 'indicate at least 1 processed to avoid multiple dialogs when To contains > 1 recipient
                End If
           End If
       Next Recipient
                
    End If
End If

Set objApp = Nothing
Set oAppt = Nothing

End Sub
Private Sub Application_NewMail()

'    msgbox "Mail received", vbOKCancel
           
End Sub
Private Sub myOLApp_ItemAdd(ByVal Item As Object)
' THIS IS NOT WORKING TO DATE
 Dim dt As Date
 Dim tm As String
 If TypeOf Item Is Outlook.MailItem And Item.SenderEmailAddress = "me@foobar.com" Then
 dt = DateAdd("d", 1, Date)
 tm = CStr(dt) & " 08:00"
 'mark = olMarkTomorrow
 Item.MarkAsTask olMarkTomorrow
 Item.FlagDueBy = tm 'Date and Time due
 Item.ReminderTime = tm 'reminder
 Item.MarkAsTask mark    'mark as task ''''edit''''
 Item.ToDoTaskOrdinal = tm 'task
 Item.TaskDueDate = dt 'Due date
 Item.TaskStartDate = dt 'Start date
 Item.UnRead = False 'mark Mail as read
 Item.Save 'save new settings
 End If
End Sub


