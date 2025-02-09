Attribute VB_Name = "動静一覧_年度切替"
Option Explicit

Sub 年度切替曜日()

    ' 2行目 ～ 22行目まで曜日を書き込む処理をループ
    Dim i As Long
    Dim Jyoubi As Date ' Julyの曜日
    
    For i = 2 To 22 Step 2 'Step 2で偶数列のみ処理
        ' 日付から曜日を取得
        Jyoubi = DateSerial(動静一覧_年度切替Form.TextBox1.Value, Cells(2, 2).Value, Cells(3, i).Value)
        
        ' 曜日をセルに書き込む
        Cells(4, i).Value = Format(Weekday(Jyoubi), "aaa")
    Next i
    
    ' 24行目から84行目まで曜日を書き込む処理をループ
    Dim j As Long
    Dim Ayoubi As Date ' Augustの曜日
    
    
    For j = 24 To 84 Step 2 'Step 2で偶数列のみ処理
        ' 日付から曜日を取得
        Ayoubi = DateSerial(動静一覧_年度切替Form.TextBox1.Value, Cells(2, 24).Value, Cells(3, j).Value)
        
        ' 曜日をセルに書き込む
        Cells(4, j).Value = Format(Weekday(Ayoubi), "aaa")
    Next j

End Sub
