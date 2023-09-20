Option Explicit
Dim strName, intLength, strMessage

' Prompt the user for their name
strName = InputBox("Enter your name:", "VBScript Greeting")

' Check if the name was provided
If strName <> "" Then
    ' Get the length of the name
    intLength = Len(strName)

    ' Create a personalized message based on the name's length
    If intLength <= 4 Then
        strMessage = "That's a short and sweet name, " & strName & "!"
    ElseIf intLength >= 10 Then
        strMessage = "You have a majestic long name, " & strName & "!"
    Else
        strMessage = "Hello, " & strName & "! Nice to meet you."
    End If

    ' Display the personalized greeting
    MsgBox strMessage, vbInformation, "Greeting"
Else
    ' Display a message if the user didn't provide a name
    MsgBox "You didn't enter a name!", vbExclamation, "Oops"
End If
