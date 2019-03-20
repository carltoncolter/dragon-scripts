'Command Name: insert signature
'Description: Inserts the first signature
'Group: Outlook
'---------------------------------------------------------------------------------
'Availability: Application Specific
'Application: Microsoft Outlook
'---------------------------------------------------------------------------------
'Command Type: Advaced Scripting

Declare Function keybd_event Lib "user32.dll" (ByVal vKey As _
Long, bScan As Long, ByVal Flag As Long, ByVal exInfo As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const VK_RETURN = 13
Const VK_SHIFT = 16
Const VK_MENU = 18
Const VK_LWIN = 91

Sub SendString (KeyText As String)
	Dim i As Integer

	For i = 1 To Len(KeyText)
		PressChar (Mid(KeyText,i,1))
	Next
End Sub

Sub ShiftDown
	keybd_event (VK_SHIFT, 0, 0, 0 )
End Sub

Sub ShiftUp
	keybd_event (VK_SHIFT, 0, 2, 0 )
End Sub


Sub PressChar (char As String, Optional ByVal ms As Integer = 10)
    Dim ascii As Integer
	Dim shift As Boolean
	Dim keycode As Integer

    shift = False
    char = Mid(char,1,1)
    ascii = AscW(Mid(char,1,1))

    If (ascii >=65 And ascii <=90) Then
        keycode = ascii
	    shift = True
	ElseIf (ascii >=97 And ascii <= 122) Then
		keycode = ascii - 32
    ElseIf (InStr(1, "!@#$%^&*()", char) <> 0) Then
        keycode = AscW(Mid("1234567890", InStr(1, "!@#$%^&*()", char), 1))
        shift = True
    ElseIf (ascii >=48 And ascii <=57) Then
        keycode = ascii
    Else
		Select Case char
			Case "~"
				keycode = 192
				shift = True
			Case "`"
				keycode = 192
			Case "_"
				keycode = 189
                shift = True
			Case "-"
				keycode = 189
			Case "="
				keycode = 187
			Case "+"
				keycode = 187
                shift = True
			Case "["
				keycode = 219
			Case "{"
				keycode = 219
                shift = True
            Case "]"
                keycode = 221
            Case "}"
                keycode = 221
                shift = True
            Case "\"
                keycode = 220
            Case "|"
                keycode = 220
                shift = True
            Case ";"
                keycode = 186
            Case ":"
                keycode = 186
                shift = True
            Case "'"
                keycode = 222
            Case Chr(34)
                keycode = 222
                shift = True
            Case ","
                keycode = 188
            Case "<"
                keycode= 188
                shift = True
            Case "."
                keycode = 190
            Case ">"
                keycode = 190
                shift = True
            Case "/"
                keycode = 191
            Case "?"
                keycode = 191
                shift = True
            Case Chr(9)
                keycode = 9
        End Select
    End If

    If (shift=True) Then
        ShiftDown
    End If

    PressKey (keycode, ms)

    If (shift=True) Then
        ShiftUp
    End If
End Sub

Sub PressKey (ByVal key As Integer, Optional ByVal ms As Integer = 10)
	keybd_event (key, 0, 0, 0 )
	Sleep ms
	keybd_event (key, 0, 2, 0 )
End Sub

Sub SendAlt
	PressKey VK_MENU, 500
End Sub

Sub SendEnter
	PressKey VK_RETURN, 500
End Sub

Sub Main
    SendString "Sincerely,"
	SendEnter
	SendAlt
	SendString "E2AS"
	'SendSystemKeys ("E2AS")
	SendEnter
End Sub
