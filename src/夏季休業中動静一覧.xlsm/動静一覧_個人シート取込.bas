Attribute VB_Name = "���Èꗗ_�l�V�[�g�捞"
Option Explicit

Sub �f�[�^��荞��()

    Application.ScreenUpdating = False

    Dim i As Long
    Dim r As Long
    Dim cnt As Long
    Dim file
    Dim filebook
    Dim dousei As Workbook

    r = 6
    cnt = 2
    
    ' �����t�@�C����I������_�C�A���O��\��
    file = Application.GetOpenFilename(MultiSelect:=True)
    
    ' �����t�@�C����I�������ꍇ�A�z�񂪕Ԃ����̂ŁA�����𑱂���
    If IsArray(file) Then
        ' �I�����ꂽ�e�t�@�C���ɑ΂��ď������s��
        For Each filebook In file
            ' �I�������t�@�C�����J��
            Workbooks.Open (filebook)

            ' �J�����t�@�C����dousei�Ƃ��Đݒ�
            Set dousei = ActiveWorkbook

            ' �w�肳�ꂽ�V�[�g�̒l���R�s�[
            Workbooks("�ċG�x�ƒ����Èꗗ.xlsm").Sheets("���Õ\�ꗗ").Range("A" & r).Value = dousei.Sheets("���Õ\").Range("O2").Value
            
            ' ���Õ\ �������ݗp�V�[�g��Q��̃f�[�^��ǂݍ���
            For i = 5 To dousei.Sheets("���Õ\").Range("Q10000").End(xlUp).Row
                ' �f�[�^�𓮐Õ\�ɓ\��t��
                Workbooks("�ċG�x�ƒ����Èꗗ.xlsm").Sheets("���Õ\�ꗗ").Cells(r, cnt).Value = dousei.Sheets("���Õ\").Range("Q" & i).Value
                Workbooks("�ċG�x�ƒ����Èꗗ.xlsm").Sheets("���Õ\�ꗗ").Cells(r, cnt + 1).Value = dousei.Sheets(1).Range("R" & i).Value

                ' ��ԍ������ɐi�߂�
                cnt = cnt + 2
            Next i

            dousei.Close
            
            ' ���̍s�Ɨ�ԍ��̐ݒ�
            r = r + 1
            cnt = 2
            
        Next
        MsgBox "�V�[�g�����m�ɃR�s�[����܂����B", vbInformation
    End If
    
End Sub
