Attribute VB_Name = "Module1"
Sub chatBot()
Dim userString As String
Dim user1String As String
Dim user2String As String
userString = InputBox("�п�J�ANAME")
MsgBox "�ܰ����{�ѧA:" & userString, , "NAME"

user1String = InputBox("�п�J�A���ʺ�")

MsgBox "�A���p�W�O:" & user1String, , "�ʺ�"

user2String = InputBox("�ҥH�A��ڦӨ��O")
MsgBox "�̷R�A�F:" & userString & user1String & user2String, , "���G"
End Sub
