Attribute VB_Name = "Module1"
Sub �ʺA�X��()
Dim shtIdx As Integer
For shtIdx = 2 To Sheets.Count
Sheets(shtIdx).Activate

Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S�]�w�|�̷�¾��������
Dim i, j As Long '�ŧii���̫�Aj�H�`��� I���̫�@�C J����e�C����
Dim myrng As Range
i = Cells(Rows.Count, 1).End(xlUp).Row

MsgBox "A��즳��Ƴ̫�@�C���C����" & i
For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����ASTEP_-1���˽�
Set myrng = Cells(j, "A") '�ثe�d��
If myrng = myrng.Offset(-1, 0) Then
myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
End If
Next
Application.DisplayAlerts = True '���s�}�Ҧ۫״���
Next
End Sub
