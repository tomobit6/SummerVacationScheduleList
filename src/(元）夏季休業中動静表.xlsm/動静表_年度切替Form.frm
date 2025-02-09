VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 動静表_年度切替Form 
   Caption         =   "年度切替"
   ClientHeight    =   2364
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3156
   OleObjectBlob   =   "動静表_年度切替Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "動静表_年度切替Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Range("B5:P46").Interior.ColorIndex = 0
    
    Call 年度切替曜日
    
    Unload 動静表_年度切替Form
End Sub

Private Sub UserForm_Click()
    ' 空のイベントハンドラ
End Sub
