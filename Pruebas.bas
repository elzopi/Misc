Attribute VB_Name = "Pruebas"
Private Sub pruebita()
    Dim elMensaje As Outlook.MailItem
    
    Set elMensaje = Application.ActiveInspector.CurrentItem
    
    
    If TypeOf elMensaje Is Outlook.MailItem And Len(elMensaje.Categories) = 0 Then
        'Set Item = Application.ActiveInspector.currentItem
        elMensaje.ShowCategoriesDialog
        With elMensaje
            .MarkAsTask olMarkThisWeek
      ' sets a due date in 3 days
            .TaskDueDate = Now + 5
            
            .ReminderSet = True
            .ReminderTime = Now + 5
            .Save
        End With
        
    End If
End Sub


    Dim pos As Integer
    Dim pos1 As Integer
    
    pos = InStr(1, strOrig, sStart) + 1
    pos1 = InStr(pos, strOrig, sEnd)
    stringBetween = Mid(strOrig, pos, pos1 - pos)
    

End Function

Public Function stringAfter(sIn As String, bString As String) As String
   If InStr(sIn, bString) = 0 Then
      stringAfter = ""
      Exit Function
   Else
      stringAfter = Split(sIn, bString)(1)
   End If
End Function

Public Function stringBefore(sIn As String, sEnd As String) As String
    Dim pos As Integer

    stringBefore = Split(sIn, sEnd)(0)
    


End Function

