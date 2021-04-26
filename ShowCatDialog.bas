Attribute VB_Name = "ShowCatDialog"
Public Sub ShowCatDialog()
    Dim olMessage As Outlook.MailItem
    Set olMessage = Application.ActiveInspector.CurrentItem
    olMessage.ShowCategoriesDialog
End Sub
