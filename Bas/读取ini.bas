Attribute VB_Name = "读取ini"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Function GetFromIni(ByVal strSectionHeader As String, ByVal strVariableName As String, ByVal strFilename As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = StripTerminator(Left$(strReturn, GetPrivateProfileString(strSectionHeader, strVariableName, "", strReturn, Len(strReturn), strFilename)))
End Function
Function WriteIni(ByVal strSectionHeader As String, ByVal strVariableName As String, ByVal strVariableValue As String, ByVal strFilename As String) As String
Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSectionHeader, strVariableName, strVariableValue, strFilename)
    If lngReturn = 0 Then
        MsgBox Load_Lanuage("文件写入失败（请检查您是否拥有在此保存的权限）", "Public", "Write_File_Fail", Lanuage_Now)
    Else
        'MsgBox   "Sucessed"
    End If
End Function
