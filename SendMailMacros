Sub addPrefix()

 Dim objApp As Outlook.Application
 Dim myItem As Outlook.MailItem
 
 Set objApp = Application
 Set myItem = objApp.ActiveInspector.CurrentItem
 
 myItem.Body = Replace(myItem.Body, vbCrLf, vbCrLf & ">")
 
 Set objApp = Nothing
    
End Sub

