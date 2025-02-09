Attribute VB_Name = "動静一覧_列数確認"
Option Explicit

Sub コード確認用メッセージボックス()
    MsgBox Cells(5, Columns.Count).End(xlToLeft).Column
End Sub
