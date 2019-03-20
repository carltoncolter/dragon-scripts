'Command Name: <move_or_dock> window <direction>
'Description: Moves or docks the window to a particular direction.  This is
'             helpful when you need to move a window to a different monitor.
'Group: Windows
'---------------------------------------------------------------------------------
'Availability: Global
'---------------------------------------------------------------------------------
'Command Type: Advaced Scripting

'keybd_event from user32.dll
Declare Function keybd_event Lib "user32.dll" (ByVal vKey As _
Long, bScan As Long, ByVal Flag As Long, ByVal exInfo As Long) As Long

'Virtual Keys
Const VK_SHIFT = 16
Const VK_LWIN = 91

Sub Main
    Dim Direction As String
    Direction = UtilityProvider.ContextValue(1)
    Direction = UCase(Left(Direction, 1)) & Mid(Direction, 2)
    If UtilityProvider.ContextValue(0)= "move" Then
	    Direction = "{Shift+" & Direction & "}"
    Else
	    Direction = "{" & Direction & "}"
    End If
    
    'Left Windows Key down
    keybd_event(VK_LWIN,0,0,0)
    'Send key codes to move window
    SendSystemKeys Direction
    'Left Windows Key up
    keybd_event(VK_LWIN,0,2,0)

    'Clear direction
    Direction = ""
End Sub