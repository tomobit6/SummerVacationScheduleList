Attribute VB_Name = "���Èꗗ_�N�x�ؑ�"
Option Explicit

Sub �N�x�ؑ֗j��()

    ' 2�s�� �` 22�s�ڂ܂ŗj�����������ޏ��������[�v
    Dim i As Long
    Dim Jyoubi As Date ' July�̗j��
    
    For i = 2 To 22 Step 2 'Step 2�ŋ�����̂ݏ���
        ' ���t����j�����擾
        Jyoubi = DateSerial(���Èꗗ_�N�x�ؑ�Form.TextBox1.Value, Cells(2, 2).Value, Cells(3, i).Value)
        
        ' �j�����Z���ɏ�������
        Cells(4, i).Value = Format(Weekday(Jyoubi), "aaa")
    Next i
    
    ' 24�s�ڂ���84�s�ڂ܂ŗj�����������ޏ��������[�v
    Dim j As Long
    Dim Ayoubi As Date ' August�̗j��
    
    
    For j = 24 To 84 Step 2 'Step 2�ŋ�����̂ݏ���
        ' ���t����j�����擾
        Ayoubi = DateSerial(���Èꗗ_�N�x�ؑ�Form.TextBox1.Value, Cells(2, 24).Value, Cells(3, j).Value)
        
        ' �j�����Z���ɏ�������
        Cells(4, j).Value = Format(Weekday(Ayoubi), "aaa")
    Next j

End Sub
