'Command Name: flag email|message|it
'Description: Flags on an outlook item
'Group: Outlook
'---------------------------------------------------------------------------------
'Availability: Application Specific
'Application: Microsoft Outlook
'---------------------------------------------------------------------------------
'Command Type: Advaced Scripting

Const CATEGORY1 As String = "1. Important and Urgent"
Const CATEGORY2 As String = "2. Important but Not Urgent"
Const CATEGORY3 As String = "3. Not Important but Urgent"
Const CATEGORY4 As String = "4. Not Important and Not Urgent"
Const CATEGORY5 As String = "5. Spare Time"


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

'! Gets the flag category:
'!    1. Importand and Urgent            (RED)
'!    2. Importand but Not Urgent        (ORANGE)
'!    3. Not Important but Urgent        (YELLOW)
'!    4. Not Important and Not Urgent    (GREEN)
'!    5. Spare Time                      (BLUE)
'!
'! @param  level     The category level number
'! @return a string value for the updated category field
Function GetFlagCategory(level As Integer)
    Dim Category As String
    Select Case level
        Case 1:
            Category = CATEGORY1
        Case 2:
            Category = CATEGORY2
        Case 3:
            Category = CATEGORY3
        Case 4:
            Category = CATEGORY4
        Case 5:
            Category = CATEGORY5
    End Select

    GetFlagCategory = Category
End Function

'! Appends a category to a string.
'! 
'! @param  objMsg    The outlook item
'! @param  category  The category to add
'! @return a string value for the updated category field
Function AppendCategory(objMsg as Object, category As String)
    Dim arr
    Dim cat As String
    Dim newCat As String
    newCat = ""

    Dim j As Integer
    j = 0

    arr = Split(objMsg.Categories, ",")
    If UBound(arr) >= 0 Then
    ' Check for Category
        For i = 0 To UBound(arr)
            cat = Trim(arr(i))
            If (cat <> CATEGORY1 And cat <> CATEGORY2 And cat <> CATEGORY3 And cat <> CATEGORY4 And cat <> CATEGORY5) Then
                If j <> 0 Then
                    newCat = newCat & ", " & cat
                Else
                    newCat = cat
                End If
                j = 1
        End If
        Next
    End If

    If Len(newCat) > 0 Then
        newCat = category & ", " & newCat
    Else
        newCat = Category
    End If

    AppendCategory = newCat
End function

'! Gets the flag color:
'!    1. Importand and Urgent            (RED)
'!    2. Importand but Not Urgent        (ORANGE)
'!    3. Not Important but Urgent        (YELLOW)
'!    4. Not Important and Not Urgent    (GREEN)
'!    5. Spare Time                      (BLUE)
'!
'! @param  objMsg    The outlook item
'! @return an integer value for the flag color
Function GetFlagColor(level As Integer)
    Const FLAGCOLOR1 As Integer = olRedFlagIcon
    Const FLAGCOLOR2 As Integer = olOrangeFlagIcon
    Const FLAGCOLOR3 As Integer = olYellowFlagIcon
    Const FLAGCOLOR4 As Integer = olGreenFlagIcon
    Const FLAGCOLOR5 As Integer = olBlueFlagIcon

    Dim FLAGCOLOR As Integer

    Select Case level
        Case 1:
            FLAGCOLOR = FLAGCOLOR1
        Case 2:
            FLAGCOLOR = FLAGCOLOR2
        Case 3:
            FLAGCOLOR = FLAGCOLOR3
        Case 4:
            FLAGCOLOR = FLAGCOLOR4
        Case 5:
            FLAGCOLOR = FLAGCOLOR5
    End Select

    GetFlagColor = FLAGCOLOR
End Function


'! Gets the subject of an outlook item.
'!
'! @param  objMsg    The outlook item
'! @return the subject of a message without any FW or RE tags
Function GetSubject(objMsg As Object)
    Dim subject As String
    subject = objMsg.Subject
    subject = Replace(subject, "FW: ", "")
    subject = Replace(subject, "FW:", "")
    subject = Replace(subject, "RE: ", "")
    subject = Replace(subject, "RE:", "")
    GetSubject = subject
End Function

'! Flags an outlook item (objMsg)
'!
'! @param  objMsg    The outlook item
Sub FlagMessage(objMsg As Object, category As String, flagColor As Integer, subject As String)

    Const USERPROP_FOLLOWUP_DATE As String = "AddFollowUpDate"

    If (IsNull(objMsg)) Then
        'Nothing to do because there is no item.
    Else
        With objMsg
            .Categories = AppendCategory(objMsg, category)
            .FlagIcon = flagColor
            .FlagRequest = "Follow up (Priority " & category & ") - " & subject
            .FlagStatus = olFlagMarked
            .Save

            Set up = .UserProperties.Find(USERPROP_FOLLOWUP_DATE)

            If Not up Is Nothing Then
                .MarkAsTask olMarkNoDate
                .Save
            End If
        End With
    End If
End Sub

'! Converts string word to number
'!
'! @param  num    The number string
'! @return the number as an integer
Function str2int (num As String )
    Select Case num
        Case "one"
            str2int = 1
        Case "two"
            str2int = 2
        Case "three"
            str2int = 3
        Case "four"
            str2int = 4
        Case "five"
            str2int = 5
        Case "1"
            str2int = 1
        Case "2"
            str2int = 2
        Case "3"
            str2int = 3
        Case "4"
            str2int = 4
        Case "5"
            str2int = 5
    End Select
End Function

Sub Main()
    Dim outlookItem As Object
    Set outlookItem = GetCurrentItem()
    
    Dim categoryLevel As Integer
    categoryLevel = str2int(UtilityProvider.ContextValue(0))

    Dim flagColor As Integer
    flagColor = GetFlagColor(categoryLevel)

    Dim category As String
    category = GetFlagCategory(categoryLevel)
    
    Dim subject As String
    subject = GetSubject(outlookItem)
    
    FlagMessage (outlookItem, category, flagColor, subject)
    Set outlookItem = Nothing
End Sub