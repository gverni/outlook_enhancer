'**********************************
'**** SNOOZE EMAIL FUNCTIONS ******
'**********************************

'*** Declarations ***
Option Explicit
Option Compare Text

'*** Functions ***

Sub MoveAllMailsToFolder(srcFolderName As String, destFolderName As String)

    On Error Resume Next
        
    Dim objSrcFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace
    Dim objItem As Object
              
    ' Get Inbox
    Set objNS = Application.GetNamespace("MAPI")
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
    
    'Get src folders
    Set objSrcFolder = objInbox.Folders(srcFolderName)
    
    
    For i = objSrcFolder.Items.Count To 1 Step -1
        objSrcFolder.Items(i).UnRead = True
        objSrcFolder.Items(i).Move objInbox.Folders("@" & objSrcFolder.Items(i).Categories)
    Next i

    Set objItem = Nothing
    Set objDestFolder = Nothing
    Set objSrcFolder = Nothing
    Set objNS = Nothing

End Sub

'Private Sub Application_Quit()
'  If TimerID <> 0 Then Call DeactivateTimer 'Turn off timer upon quitting **VERY IMPORTANT**
'End Sub
'
'Private Sub Application_Startup()
'  Call ActivateTimer(30) 'Set timer to go off every 30 minutes
'End Sub
'
'
Public Sub ActivateTimer(ByVal nMinutes As Long)
  nMinutes = nMinutes * 1000 * 60 'The SetTimer call accepts milliseconds, so convert to minutes
  If TimerID <> 0 Then Call DeactivateTimer 'Check to see if timer is running before call to SetTimer
  TimerID = SetTimer(0, 0, nMinutes, AddressOf TriggerTimer)
  If TimerID = 0 Then
    MsgBox "The timer failed to activate."
  End If
End Sub

Public Sub DeactivateTimer()
Dim lSuccess As Long
  lSuccess = KillTimer(0, TimerID)
  If lSuccess = 0 Then
    MsgBox "The timer failed to deactivate."
  Else
    TimerID = 0
  End If
End Sub

Public Sub TriggerTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)

    MoveAllMailsToFolder "@Snoozed", ""

End Sub
