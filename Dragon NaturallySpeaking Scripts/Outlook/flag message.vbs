'Command Name: flag email|message|it
'Description: Flags on an outlook item
'Group: Outlook
'---------------------------------------------------------------------------------
'Availability: Application Specific
'Application: Microsoft Outlook
'---------------------------------------------------------------------------------
'Command Type: Advaced Scripting

'! Gets the current outlook item.  Either the item selected in a list or the
'! currently open item.
'!
'! @return The selected outlook item object.
'! @see http://slipstick.me/e8mio for the GetCurrentItem function
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select

    Set objApp = Nothing
End Function

'! Flags an outlook item (objMsg)
'!
'! @param  objMsg    The outlook item
Sub FlagMessage(objMsg As Object)

    Const USERPROP_FOLLOWUP_DATE As String = "AddFollowUpDate"

    If (IsNull(objMsg)) Then
        'Nothing to do because there is no item.
    Else
        With objMsg
            ' Set the Follow up Flag
            .FlagIcon = olRedFlagIcon
            .FlagRequest = "Follow up"
            .FlagStatus = olFlagMarked
            ' Set the due date for the reminder two days from today
            .FlagDueBy = DateAdd("d", 1, Date)
            .Save

            Set up = .UserProperties.Find(USERPROP_FOLLOWUP_DATE)

            If Not up Is Nothing Then
                .MarkAsTask olMarkNoDate
                .TaskDueDate = up.Value
                .ReminderSet = True
                .ReminderTime = up.Value
                .Save
            End If

        End With
    End If
End Sub

Sub Main()
    Dim outlookItem As Object
    Set outlookItem = GetCurrentItem()
    FlagMessage (outlookItem)
    Set outlookItem = Nothing
End Sub