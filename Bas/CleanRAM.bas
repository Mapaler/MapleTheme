Attribute VB_Name = "�ڴ�����"
'/* �ڴ��������� */
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'/** �ڴ��Զ����� **/
Public Sub NeiCun_Timer()
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1
End Sub
