VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���Èꗗ_�N�x�ؑ�Form 
   Caption         =   "�N�x�ؑ�"
   ClientHeight    =   2088
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   2772
   OleObjectBlob   =   "���Èꗗ_�N�x�ؑ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���Èꗗ_�N�x�ؑ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Range("B3:CG4,B6:CG20").Interior.ColorIndex = 0
    Range("B5:CG5").Interior.Color = RGB(155, 194, 230)
    Range("A6:CG20").ClearContents
    
    Call �N�x�ؑ֗j��
    
    Unload ���Èꗗ_�N�x�ؑ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃C�x���g�n���h��
End Sub
