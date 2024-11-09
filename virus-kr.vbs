' github.com/anion15
Dim userResponse
Do
userResponse = MsgBox( _
    "안녕하세요, 저는 바이러스입니다." & vbCrLf & _
    "하지만 기술이 부족해서 불행히도 귀하의 컴퓨터에 해를 끼칠 수 없습니다." & vbCrLf & _
    "귀하의 중요한 파일 중 하나를 직접 삭제한 다음 다른 사용자에게 전달해 주시기를 바랍니다." & vbCrLf & _
    "협조해 주셔서 감사합니다!", _
    vbYesNoCancel + vbExclamation, "바이러스 경고!")

    If userResponse = vbYes Then
        Exit Do
    End If
    If userResponse = vbNo Or userResponse = vbCancel Then
        Exit Do
    End If
    
Loop Until userResponse = vbYes Or userResponse = vbNo Or userResponse = vbCancel
