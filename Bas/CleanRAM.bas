Attribute VB_Name = "内存清理"
'/* 内存清理声明 */
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'/** 内存自动清理 **/
Public Sub NeiCun_Timer()
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1
End Sub
