Attribute VB_Name = "動静表_年度切替"
Option Explicit

Sub 年度切替曜日()

    Cells(1, 4).Value = 動静表_年度切替Form.TextBox1.Value

    Dim i As Long
    Dim Jyoubi As Date ' Julyの曜日
    
    For i = 5 To 15
        Jyoubi = DateSerial(Cells(1, 4).Value, Cells(5, 1).Value, Cells(i, 2).Value)
        Cells(i, 3).Value = Format(Weekday(Jyoubi), "aaa")
    Next i
    
    Dim j As Long
    Dim Ayoubi As Date ' Augustの曜日
    
    For j = 16 To 46
        Ayoubi = DateSerial(Cells(1, 4).Value, Cells(16, 1).Value, Cells(j, 2).Value)
        Cells(j, 3).Value = Format(Weekday(Ayoubi), "aaa")
    Next j
    
End Sub
