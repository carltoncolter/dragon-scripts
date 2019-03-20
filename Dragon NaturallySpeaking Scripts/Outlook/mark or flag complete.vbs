'Command Name: <mark_or_flag> complete
'Description: Marks or Flags an outlook item as complete 
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

'! Marks an outlook item (objMsg) as Complete
'!
'! @param  objMsg    The outlook item
Sub FlagComplete(objMsg As Object)
    If (IsNull(objMsg)) Then
        'Nothing to do because there is no item.
    Else
        With objMsg
            ' Set the Follow up Flag
            If .FlagRequest Like "Follow up*" Or .FlagStatus <> olNoFlag Then
                .FlagStatus = olFlagComplete
            End If
            .Save
        End With
    End If
End Sub

Sub Main()
    Dim outlookItem As Object
    Set outlookItem = GetCurrentItem()
    FlagComplete (outlookItem)
    Set outlookItem = Nothing
End Sub