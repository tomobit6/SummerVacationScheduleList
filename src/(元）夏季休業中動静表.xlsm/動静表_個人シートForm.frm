VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���Õ\_�l�V�[�gForm 
   Caption         =   "�l���Õ\�쐬"
   ClientHeight    =   8088
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10752
   OleObjectBlob   =   "���Õ\_�l�V�[�gForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���Õ\_�l�V�[�gForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click() 'OK�{�^��
    Dim i As Long
    Dim text As String
    
    For i = 1 To 40
        ' �X�N���[���̍X�V���~�i�������������j
        Application.ScreenUpdating = False
        
        ' �e�L�X�g�{�b�N�X�ɓ��͂�����ꍇ�A���e��ϐ��ɑ��
        If Me.Controls("TextBox" & i).Value <> "" Then
            text = Me.Controls("TextBox" & i).Value ' ME��UserForm���w��
            
            ThisWorkbook.Worksheets("���Õ\").Copy
            
            ' �R�s�[�����V�[�g��V�����t�@�C�����ŕۑ�
            ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\�ċG�x�ƒ����Õ\(" & text & ").xlsx"
            
            ' �V�����쐬�����u�b�N�����
            ActiveWorkbook.Close False
        End If
    Next i
    
    MsgBox "�V�[�g�����m�ɃR�s�[����܂����B", vbInformation
    Unload ���Õ\_�l�V�[�gForm
End Sub

Private Sub CommandButton2_Click() '�L�����Z���{�^��
    Unload ���Õ\_�l�V�[�gForm
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()
    ' ��̃C�x���g�n���h��
End Sub
