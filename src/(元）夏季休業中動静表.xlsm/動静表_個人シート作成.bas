Attribute VB_Name = "動静表_個人シート作成"
Option Explicit

Sub 個人シート作成()
    ThisWorkbook.Worksheets("動静表").Copy
    
    ' テキストボックスの名前を取得してファイル名を作成
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\夏季休業中動静表(" & 動静表_個人シートForm.Controls("TextBox" & i).Value & ").xlsx"
End Sub



