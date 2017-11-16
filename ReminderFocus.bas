Option Explicit
Option Compare Text

'*** Declarations ***
Public remindersAlertEnabled As Boolean
Dim AlertSnoozedAt As Date
' If you have more than one appointment snoozed, this macro will fire an alert for each one of them
' To avoid that,  we use the following constant. This basically disable the alert for a certain amount of seconds
Const SNOOZEDELTA = 30

'*** Functions ***

Public Sub disableRemindersAlert()

    remindersAlertEnabled = False

End Sub

Public Sub enableRemindersAlert()

    remindersAlertEnabled = True

End Sub

Private Function isWorkingTime()
   
    isWorkingTime = Hour(Time()) >= 9 And (Hour(Time()) < 18)

End Function

Public Sub ProcessReminder(Item As Object)

    If TypeOf Item Is AppointmentItem Then
        'show message box for first reminder
        If DateDiff("s", AlertSnoozedAt, Now()) > SNOOZEDELTA And (isWorkingTime()) Then
            MsgBox "You have reminder(s)", vbSystemModal, ""
            AlertSnoozedAt = Now()
            
        End If
        
        'Below code doesn't work. Maybe one day....
        'ReminderWindow = FindWindowA("#32770", vbNullString)
        'SetWindowPos ReminderWindow, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        
        Call gverni.CategorizeInbox
        
    End If

End Sub

