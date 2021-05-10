Attribute VB_Name = "Module1"
Sub 動態合併()
Dim shtIdx As Integer
For shtIdx = 2 To Sheets.Count
Sheets(shtIdx).Activate

Application.DisplayAlerts = False '作業系統提醒文字，若沒設定會依照職提醒提醒
Dim i, j As Long '宣告i為最後，j違常整數 I為最後一列 J為當前列索引
Dim myrng As Range
i = Cells(Rows.Count, 1).End(xlUp).Row

MsgBox "A欄位有資料最後一列的列索引" & i
For j = i To 2 Step -1 '從最後一列到第二列遞減，STEP_-1為倒豎
Set myrng = Cells(j, "A") '目前範圍
If myrng = myrng.Offset(-1, 0) Then
myrng.Offset(-1, 0).Resize(2, 1).Merge '則需由下而上合併
End If
Next
Application.DisplayAlerts = True '重新開啟自度提醒
Next
End Sub
