VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���Õ\_�N�x�ؑ�Form 
   Caption         =   "�N�x�ؑ�"
   ClientHeight    =   2364
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3156
   OleObjectBlob   =   "���Õ\_�N�x�ؑ�Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���Õ\_�N�x�ؑ�Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Range("B5:P46").Interior.ColorIndex = 0
    
    Call �N�x�ؑ֗j��
    
    Unload ���Õ\_�N�x�ؑ�Form
End Sub

Private Sub UserForm_Click()
    ' ��̃C�x���g�n���h��
End Sub
