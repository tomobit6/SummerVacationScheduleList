Attribute VB_Name = "���Õ\_�N�x�ؑ�"
Option Explicit

Sub �N�x�ؑ֗j��()

    Cells(1, 4).Value = ���Õ\_�N�x�ؑ�Form.TextBox1.Value

    Dim i As Long
    Dim Jyoubi As Date ' July�̗j��
    
    For i = 5 To 15
        Jyoubi = DateSerial(Cells(1, 4).Value, Cells(5, 1).Value, Cells(i, 2).Value)
        Cells(i, 3).Value = Format(Weekday(Jyoubi), "aaa")
    Next i
    
    Dim j As Long
    Dim Ayoubi As Date ' August�̗j��
    
    For j = 16 To 46
        Ayoubi = DateSerial(Cells(1, 4).Value, Cells(16, 1).Value, Cells(j, 2).Value)
        Cells(j, 3).Value = Format(Weekday(Ayoubi), "aaa")
    Next j
    
End Sub
