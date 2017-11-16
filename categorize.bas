Option Explicit
Option Compare Text

'*** Declaration ***
Public dbConversations As Collection

'*** Functions ***

Sub moveToFolder(folderName, ByVal olCurrMailItem As MailItem)

 Dim mailboxNameString As String
 
 mailboxNameString = "gverni@qti.qualcomm.com"
 
 Dim olApp As New Outlook.Application
 Dim olNameSpace As Outlook.NameSpace
 Dim olCurrExplorer As Outlook.Explorer
 Dim olCurrSelection As Outlook.Selection
  
 Dim olDestFolder As Outlook.MAPIFolder
 'Dim olCurrMailItem As MailItem
 Dim m As Integer

 Set olNameSpace = olApp.GetNamespace("MAPI")
 Set olCurrExplorer = olApp.ActiveExplorer
 Set olCurrSelection = olCurrExplorer.Selection

 Set olDestFolder = olNameSpace.Folders(mailboxNameString).Folders("Inbox").Folders(folderName)

 olCurrMailItem.Move olDestFolder
 Debug.Print "End of MoveToFolder " & olDestFolder

End Sub


Public Function CollectionContains(col As Collection, key As Variant) As Boolean

Dim obj As Variant

On Error GoTo err
    CollectionContains = True
    obj = col(key)
    Exit Function
    
err:
    CollectionContains = False

End Function


Sub ProcessEmailForCategories(Item As Object)

On Error Resume Next
  'Item.BodyFormat = olFormatPlain
  'Item.Save
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim numRecipientsNoMl As Integer
    Dim Prio, newPrio As Integer
        
    
    Const PR_SMTP_ADDRESS As String = _
        "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    Const MyName = ""
    Const MyMls = ""
    Const prioCaseCr = -1
    Const prioAdmin = -2
    Const prioDevices = -3
    Const prioHR = -4
    Const NUM_RECIPIENTS = 10
    
    Dim From_Prio1, From_Prio2, From_Prio3, From_Prio4, from_Admin, from_HR, From_CaseCr, To_Prio1, To_Prio2, To_Prio3, To_Prio4, To_Devices, CC_Prio1, CC_Prio2, CC_Prio3, CC_Prio4, CC_Devices As String
    Dim Subject_Prio1(), Subject_Prio2(), Subject_Prio3(), Subject_Prio4() As String
    Dim strSubject_Prio4 As Variant
    
    
    From_Prio1 = ""
    From_Prio2 = ""
    From_Prio3 = ""
    From_Prio4 = ""
    from_Admin = ""
    from_HR = ""
    From_CaseCr = ""
    
    To_Prio1 = MyName & ""
    To_Prio2 = MyMls
    To_Prio3 = ""
    To_Prio4 = ""
    To_Devices = ""
    
    CC_Prio1 = ""
    CC_Prio2 = MyName
    CC_Prio3 = MyMls
    CC_Prio4 = ""
    CC_Devices = ""
    
    'Subject_Prio1() = [""]
    'Subject_Prio2() = [""]
    'Subject_Prio3() = [""]
    Subject_Prio4() = [""]

    'Set lower priority to start with
    Prio = 4
            
    'Check the sender
    Debug.Print "Processing " & Item.Sender
    If InStr(From_Prio1, Item.Sender) <> 0 Then
        Prio = 1
    ElseIf InStr(From_Prio2, Item.Sender) <> 0 Then
        Prio = 2
    ElseIf InStr(From_Prio3, Item.Sender) <> 0 Then
        Prio = 3
    ElseIf InStr(From_Prio4, Item.Sender) <> 0 Then
        Prio = 4
    ElseIf InStr(From_CaseCr, Item.Sender) <> 0 Then
        Prio = prioCaseCr
    ElseIf InStr(from_Admin, Item.Sender) <> 0 Then
        Prio = prioAdmin
    ElseIf InStr(from_HR, Item.Sender) <> 0 Then
        Prio = prioHR
    End If
    
    Debug.Print "Prio set to " & Prio
    
    'Check the To and CC
    If Prio > 1 Then 'Skip processing if the prio is negative (restrictive rules) or is already 1
        Debug.Print "Processing To and CC"
        numRecipientsNoMl = 0
        Set recips = Item.Recipients
        For Each recip In recips
            numRecipientsNoMl = numRecipientsNoMl + 1
            Select Case recip.Type
                Case OlMailRecipientType.olTo:
                    If InStr(To_Prio1, recip) <> 0 Then
                        Prio = 1
                    ElseIf InStr(To_Prio2, recip) <> 0 Then
                        If Prio > 2 Then Prio = 2
                    ElseIf InStr(To_Prio3, recip) <> 0 Then
                        If Prio > 3 Then Prio = 3
                    ElseIf InStr(To_Prio4, recip) <> 0 Then
                        If Prio > 4 Then Prio = 4
                    ElseIf InStr(To_Devices, recip) <> 0 Then
                        Prio = prioDevices
                        Exit For 'We don't need to process further
                    End If
                Case OlMailRecipientType.olCC
                    If InStr(CC_Prio1, recip) <> 0 Then
                        Prio = 1
                    ElseIf InStr(CC_Prio2, recip) <> 0 Then
                        If Prio > 2 Then Prio = 2
                    ElseIf InStr(CC_Prio3, recip) <> 0 Then
                        If Prio > 3 Then Prio = 3
                    ElseIf InStr(CC_Prio4, recip) <> 0 Then
                        If Prio > 4 Then Prio = 4
                    ElseIf InStr(CC_Devices, recip) <> 0 Then
                        Prio = prioDevices
                        Exit For 'We don't need to process further
                    End If
            End Select
        Next recip
    End If

    Debug.Print "Prio set to " & Prio
    

'The subject set prio is overriding the priority
    For Each strSubject_Prio4 In Subject_Prio4
        If InStr(Item.Subject, strSubject_Prio4) <> 0 Then
            Prio = 4
            Debug.Print "Subject overriding priority"
        End If
    Next strSubject_Prio4

'Finally, the conversation ID will override all the priorities
    If CollectionContains(dbConversations, Item.ConversationID) Then
        newPrio = dbConversations(Item.ConversationID)
        If newPrio < Prio Then
            Prio = newPrio
        End If
    Else
        dbConversations.Add Prio, Item.ConversationID
    End If
    

'TODO: Process num recipient
'    If Item.Recipients.Count > NUM_RECIPIENTS Then
'        Prio = Prio + 1
'        Debug.Print "Add 1 for num rec"
'    End If
           
    'Set Category
    Select Case (Prio)
        Case prioAdmin
            moveToFolder "@Admin", Item
        Case prioCaseCr
            Item.Categories = "Case&CR"
            moveToFolder "@CaseCR", Item
        Case prioHR
            moveToFolder "@HR", Item
        Case prioDevices
            moveToFolder "@Devices", Item
        Case Is > 0
            Item.Categories = "Prio" & Prio
            moveToFolder "@Prio" & Prio, Item
    End Select
    
    Item.Save
    
    Debug.Print "------"

End Sub

'This is called by the "Categorize InBox" button in the Ribbon
Public Sub CategorizeInbox()

  Dim objNS As NameSpace
  Set objNS = Application.Session
  ' instantiate objects declared WithEvents
  Dim MyInboxItems As Outlook.Items
  Set MyInboxItems = objNS.GetDefaultFolder(olFolderInbox).Items
  Dim myInboxItem As Object
  Dim nonMailItems As Integer
  Dim i As Integer
  
  For i = MyInboxItems.Count To 1 Step -1
      If MyInboxItems(i).Class = olMail Then ProcessEmailForCategories MyInboxItems(i)
  Next i
  
  'For Each myInboxItem In MyInboxItems
  '  If myInboxItem.Class = olMail Then olInboxItems_ItemAdd myInboxItem
  'Next
  
'  nonMailItems = 0
'  Do While MyInboxItems.Count > nonMailItems
'    If MyInboxItems(nonMailItems + 1).Class = olMail Then
'        olInboxItems_ItemAdd MyInboxItems(nonMailItems + 1)
'    Else
'        nonMailItems = nonMailItems + 1
'    End If
'  Loop

  
End Sub

