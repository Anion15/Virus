' github.com/anion15
Dim userResponse
Do
userResponse = MsgBox( _
    "�ȳ��ϼ���, ���� ���̷����Դϴ�." & vbCrLf & _
    "������ ����� �����ؼ� �������� ������ ��ǻ�Ϳ� �ظ� ��ĥ �� �����ϴ�." & vbCrLf & _
    "������ �߿��� ���� �� �ϳ��� ���� ������ ���� �ٸ� ����ڿ��� ������ �ֽñ⸦ �ٶ��ϴ�." & vbCrLf & _
    "������ �ּż� �����մϴ�!", _
    vbYesNoCancel + vbExclamation, "���̷��� ���!")

    If userResponse = vbYes Then
        Exit Do
    End If
    If userResponse = vbNo Or userResponse = vbCancel Then
        Exit Do
    End If
    
Loop Until userResponse = vbYes Or userResponse = vbNo Or userResponse = vbCancel
