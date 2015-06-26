Attribute VB_Name = "打开多个文件"
Option Explicit

Type DlgFileInfo
 iCount As Long
 sPath As String
 sFile() As String
End Type

'功能： 返回CommonDialog所选择的文件数量、路径和文件名
'参数说明: strFileName为CommonDialog的Filename属性
'函数类型: DlgFileInfo。这是一个自定义类型，其中iCount返回所选择文件的个数，sPath返回所选
'择文件的路径，sFile()返回所选择文件的文件名（不包括路径）
'注意事项: 该函数应在CommonDialog.ShowOpen方法后立即使用，以免当前路径被更改
Public Function GetDlgFileInfo(strFilename As String) As DlgFileInfo
 
 Dim sPath, tmpStr As String
 Dim sFile() As String
 Dim iCount As Integer
 Dim i As Integer
On Error GoTo ErrHandle
 
 sPath = CurDir()
 tmpStr = Right$(strFilename, Len(strFilename) - Len(sPath)) '将文件名与路径分离
 
 If Left$(tmpStr, 1) = Chr$(0) Then
 '选择了多个文件(分离后第一个字符为Chr$(0))
 For i = 1 To Len(tmpStr)
 If Mid$(tmpStr, i, 1) = Chr$(0) Then
 iCount = iCount + 1
 ReDim Preserve sFile(iCount)
 Else
 sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
 End If
 Next i
 Else
 '只选择了一个文件(注意：根目录下的文件名除去路径后左边没有"\"）
 iCount = 1
 ReDim Preserve sFile(iCount)
 If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
 sFile(iCount) = tmpStr
 End If
 
 GetDlgFileInfo.iCount = iCount
 ReDim GetDlgFileInfo.sFile(iCount)
 
 If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
 GetDlgFileInfo.sPath = sPath
 
 For i = 1 To iCount
 GetDlgFileInfo.sFile(i) = sFile(i)
 Next i
 
 Exit Function

ErrHandle:
' MsgBox "GetDlgFileInfo函数执行错误！（无文件的时候取消了）", vbOKOnly + vbCritical, "自定义函数错误"

End Function

