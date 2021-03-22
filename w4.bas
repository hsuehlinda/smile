Attribute VB_Name = "Module1"
Sub chatBot()
Dim userString As String
Dim user1String As String
Dim user2String As String
userString = InputBox("請輸入你NAME")
MsgBox "很高興認識你:" & userString, , "NAME"

user1String = InputBox("請輸入你的暱稱")

MsgBox "你的小名是:" & user1String, , "暱稱"

user2String = InputBox("所以你對我而言是")
MsgBox "最愛你了:" & userString & user1String & user2String, , "結果"
End Sub
