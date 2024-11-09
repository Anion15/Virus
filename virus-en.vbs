' github.com/anion15
Dim userResponse
Do
userResponse = MsgBox( _
    "Hello, I am a virus." & vbCrLf & _
    "But unfortunately they cannot harm your computer due to lack of skill." & vbCrLf & _
    "We want you to delete one of your important files yourself and then pass it on to someone else." & vbCrLf & _
    "Thank you for your cooperation!", _
    vbYesNoCancel + vbExclamation, "Virus warning!")

    If userResponse = vbYes Then
        Exit Do
    End If
    If userResponse = vbNo Or userResponse = vbCancel Then
        Exit Do
    End If
    
Loop Until userResponse = vbYes Or userResponse = vbNo Or userResponse = vbCancel
