Attribute VB_Name = "動静一覧_個人シート取込"
Option Explicit

Sub データ取り込み()

    Application.ScreenUpdating = False

    Dim i As Long
    Dim r As Long
    Dim cnt As Long
    Dim file
    Dim filebook
    Dim dousei As Workbook

    r = 6
    cnt = 2
    
    ' 複数ファイルを選択するダイアログを表示
    file = Application.GetOpenFilename(MultiSelect:=True)
    
    ' 複数ファイルを選択した場合、配列が返されるので、処理を続ける
    If IsArray(file) Then
        ' 選択された各ファイルに対して処理を行う
        For Each filebook In file
            ' 選択したファイルを開く
            Workbooks.Open (filebook)

            ' 開いたファイルをdouseiとして設定
            Set dousei = ActiveWorkbook

            ' 指定されたシートの値をコピー
            Workbooks("夏季休業中動静一覧.xlsm").Sheets("動静表一覧").Range("A" & r).Value = dousei.Sheets("動静表").Range("O2").Value
            
            ' 動静表 書き込み用シートのQ列のデータを読み込む
            For i = 5 To dousei.Sheets("動静表").Range("Q10000").End(xlUp).Row
                ' データを動静表に貼り付け
                Workbooks("夏季休業中動静一覧.xlsm").Sheets("動静表一覧").Cells(r, cnt).Value = dousei.Sheets("動静表").Range("Q" & i).Value
                Workbooks("夏季休業中動静一覧.xlsm").Sheets("動静表一覧").Cells(r, cnt + 1).Value = dousei.Sheets(1).Range("R" & i).Value

                ' 列番号を次に進める
                cnt = cnt + 2
            Next i

            dousei.Close
            
            ' 次の行と列番号の設定
            r = r + 1
            cnt = 2
            
        Next
        MsgBox "シートが正確にコピーされました。", vbInformation
    End If
    
End Sub
