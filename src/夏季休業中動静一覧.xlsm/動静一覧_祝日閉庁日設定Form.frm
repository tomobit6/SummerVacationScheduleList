VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 動静一覧_祝日閉庁日設定Form 
   Caption         =   "祝日閉庁日設定"
   ClientHeight    =   5688
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10128
   OleObjectBlob   =   "動静一覧_祝日閉庁日設定Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "動静一覧_祝日閉庁日設定Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click() 'OKボタン
    Dim i As Long
    
    For i = 1 To 42 '規則で夏休みは、毎年7月21日〜8月31日と決まっている。7月の11日＋8月の31日で42日。
        ' チェックボックスがチェックされている場合、対応するセル範囲の背景色変更
        If Me.Controls("CheckBox" & i).Value = True Then
            Range(Cells(3, 2 * i), Cells(20, 2 * i + 1)).Interior.Color = RGB(217, 225, 242)
        End If
    Next i
    
    Unload 動静一覧_祝日閉庁日設定Form
End Sub

Private Sub CommandButton2_Click() 'キャンセルボタン
    Unload 動静一覧_祝日閉庁日設定Form
End Sub

Private Sub UserForm_Click()
    ' 空のイベントハンドラ
End Sub
