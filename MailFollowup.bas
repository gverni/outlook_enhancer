Option Explicit
Option Compare Text

' This is the function called by the "Move To Follow-up" button in the Ribbon

Sub MoveLastMailToFollowUp()

On Error Resume Next

    Dim objDestFolder As Outlook.MAPIFolder, objSentFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace
    Dim objLastMail, objLastMailCopy As Object ' Mail objects
    Dim myItems As Outlook.Items
    Const destFolderName = "@followup"
    
    'Get Sent folder
    Set objNS = Application.GetNamespace("MAPI")
    Set objSentFolder = objNS.GetDefaultFolder(olFolderSentMail)
    
    'Get Destination folder. In this case is the followup folder inside the Inbox
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
    Set objDestFolder = objInbox.Folders(destFolderName)
    If objDestFolder Is Nothing Then
        MsgBox destFolderName + " folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub
    End If

    'Sort sent items based on creation time
    Set myItems = objSentFolder.Items
    myItems.Sort "[CreationTime]", True
    'Get the latest created item
    Set objLastMail = myItems.GetFirst
    ' Create a copy
    Set objLastMailCopy = objLastMail.Copy
    ' Move to destination folder
    objLastMailCopy.Move objDestFolder

    ' Destroy everything
    Set objSentFolder = Nothing
    Set objInbox = Nothing
    Set objDestFolder = Nothing
    Set objNS = Nothing
    Set objLastMail = Nothing
    Set objLastMailCopy = Nothing

End Sub


' *** This is no Longer used ***
' This is called by the "Follow-up" button in the new mail

Public Sub MarkForFollowup()

    Dim NewMail As MailItem, oInspector As Inspector
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        MsgBox "No active inspector"
    Else
        Set NewMail = oInspector.CurrentItem
        If NewMail.Sent Then
            MsgBox "This is not an editable email"
        Else
            NewMail.Categories = "Followup"
        End If
    End If
End Sub

