VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 動静表_個人シートForm 
   Caption         =   "個人動静表作成"
   ClientHeight    =   8088
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10752
   OleObjectBlob   =   "動静表_個人シートForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "動静表_個人シートForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click() 'OKボタン
    Dim i As Long
    Dim text As String
    
    For i = 1 To 40
        ' スクリーンの更新を停止（処理を高速化）
        Application.ScreenUpdating = False
        
        ' テキストボックスに入力がある場合、内容を変数に代入
        If Me.Controls("TextBox" & i).Value <> "" Then
            text = Me.Controls("TextBox" & i).Value ' MEはUserFormを指す
            
            ThisWorkbook.Worksheets("動静表").Copy
            
            ' コピーしたシートを新しいファイル名で保存
            ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\夏季休業中動静表(" & text & ").xlsx"
            
            ' 新しく作成したブックを閉じる
            ActiveWorkbook.Close False
        End If
    Next i
    
    MsgBox "シートが正確にコピーされました。", vbInformation
    Unload 動静表_個人シートForm
End Sub

Private Sub CommandButton2_Click() 'キャンセルボタン
    Unload 動静表_個人シートForm
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()
    ' 空のイベントハンドラ
End Sub
