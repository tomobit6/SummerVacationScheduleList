VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���Õ\_�j�������ݒ�Form 
   Caption         =   "�j���E�����ݒ�"
   ClientHeight    =   5724
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10068
   OleObjectBlob   =   "���Õ\_�j�������ݒ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���Õ\_�j�������ݒ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click() 'OK�{�^��
    Dim i As Long
    
    For i = 1 To 42 '�K���ŉċx�݂́A���N7��21���`8��31���ƌ��܂��Ă���B7����11���{8����31����42���B
        ' �`�F�b�N�{�b�N�X���`�F�b�N����Ă���ꍇ�A�Ή�����Z���̔w�i�F��ύX
        If Me.Controls("CheckBox" & i).Value = True Then
            Range(Cells(i + 4, 2), Cells(i + 4, 16)).Interior.Color = RGB(217, 225, 242)
        End If
    Next i
    
    Unload ���Õ\_�j�������ݒ�Form
End Sub

Private Sub CommandButton2_Click() '�L�����Z���{�^��
    Unload ���Õ\_�j�������ݒ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃C�x���g�n���h��
End Sub
