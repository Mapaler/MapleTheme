Attribute VB_Name = "�򿪶���ļ�"
Option Explicit

Type DlgFileInfo
 iCount As Long
 sPath As String
 sFile() As String
End Type

'���ܣ� ����CommonDialog��ѡ����ļ�������·�����ļ���
'����˵��: strFileNameΪCommonDialog��Filename����
'��������: DlgFileInfo������һ���Զ������ͣ�����iCount������ѡ���ļ��ĸ�����sPath������ѡ
'���ļ���·����sFile()������ѡ���ļ����ļ�����������·����
'ע������: �ú���Ӧ��CommonDialog.ShowOpen����������ʹ�ã����⵱ǰ·��������
Public Function GetDlgFileInfo(strFilename As String) As DlgFileInfo
 
 Dim sPath, tmpStr As String
 Dim sFile() As String
 Dim iCount As Integer
 Dim i As Integer
On Error GoTo ErrHandle
 
 sPath = CurDir()
 tmpStr = Right$(strFilename, Len(strFilename) - Len(sPath)) '���ļ�����·������
 
 If Left$(tmpStr, 1) = Chr$(0) Then
 'ѡ���˶���ļ�(������һ���ַ�ΪChr$(0))
 For i = 1 To Len(tmpStr)
 If Mid$(tmpStr, i, 1) = Chr$(0) Then
 iCount = iCount + 1
 ReDim Preserve sFile(iCount)
 Else
 sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
 End If
 Next i
 Else
 'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·�������û��"\"��
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
' MsgBox "GetDlgFileInfo����ִ�д��󣡣����ļ���ʱ��ȡ���ˣ�", vbOKOnly + vbCritical, "�Զ��庯������"

End Function

