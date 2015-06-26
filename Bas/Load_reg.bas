Attribute VB_Name = "读取注册表"
Option Explicit
'================================
'注册表通用操作函数
'================================


'==================================================
'注册表操作声明
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Private Const REG_SZ = 1&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&
Private Const ERROR_SUCCESS = 0&
'==================================================

'================================
'注册表操作函数
'================================


'读取注册表字符串键值
Public Function GetString(hKey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim lValueType As Long 'new add
RegOpenKey hKey, strPath, keyhand
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then
strBuf = String(lDataBufSize, " ")
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal strBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
intZeroPos = InStr(strBuf, Chr$(0))
If intZeroPos > 0 Then
GetString = StripTerminator(Left$(strBuf, intZeroPos - 1))
Else: GetString = StripTerminator(strBuf)
End If
End If
End If
End Function

'写入注册表字符串键值
Public Sub SetString(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
RegCreateKey hKey, strPath, keyhand
RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode))
RegCloseKey keyhand
End Sub

'读取注册表 DWORD 键值
Function GetDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(hKey, strPath, keyhand)

' Get length/data type
lDataBufSize = 4

lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
If lValueType = REG_DWORD Then
GetDword = lBuf
End If
'Else
' Call errlog("GetDWORD-" & strPath, False)
End If

r = RegCloseKey(keyhand)
End Function

'写入注册表 DWORD 键值
Function SetDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
Dim keyhand As Long
RegCreateKey hKey, strPath, keyhand
RegSetValueEx keyhand, strValueName, 0&, REG_DWORD, lData, 4
RegCloseKey keyhand
End Function

'读取注册表二进制键值
Function GetBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(hKey, strPath, keyhand)

' Get length/data type
lDataBufSize = 4

lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
If lValueType = REG_BINARY Then
GetBinary = lBuf
End If
End If

r = RegCloseKey(keyhand)
End Function

'写入注册表二进制键值
Function SetBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long, ByVal BitNumber As Long)
Dim keyhand As Long
RegCreateKey hKey, strPath, keyhand
RegSetValueEx keyhand, strValueName, 0&, REG_BINARY, lData, BitNumber
RegCloseKey keyhand
End Function

'删除一个注册表键值
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
RegOpenKey hKey, strPath, keyhand
RegDeleteValue keyhand, strValue
RegCloseKey keyhand
End Function

'创建一个主键
Public Function CreateKey(ByVal hKey As Long, ByVal strKey As String)
Dim keyhand&
RegCreateKey hKey, strKey, keyhand
RegCloseKey keyhand&
End Function

'自写函数：此函数用于去掉所有不必要的 Chr$(0) 中止符
Public Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    '搜索第一个 Chr$(0) 中止符
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Public Sub GetAllKey(ByVal hResult As Long, ByVal strKey As String, ByRef vName As Variant, Optional ByVal dNum As Integer = -1, Optional ByVal Num2 As Integer = 0)
Dim lngFilterIndex As Long

    Dim hKey As Long, Cnt As Long, sSave As String
    '打开一个注册表键
    RegOpenKey hResult, strKey, hKey
    '列出该键下的所有子键
    Cnt = 0

If dNum < 0 Then '一维数组
    Do
        '创建一个缓冲区
        sSave = String(255, 0)
        '枚举出所有子键
        If RegEnumKeyEx(hKey, Cnt, sSave, 255, 0, vbNullString, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        '将枚举出来的子键逐一加入数组
        ReDim Preserve vName(0 To Cnt)
        vName(Cnt) = StripTerminator(sSave)
        Cnt = Cnt + 1
    Loop
Else
    Do
        '创建一个缓冲区
        sSave = String(255, 0)
        '枚举出所有子键
        If RegEnumKeyEx(hKey, Cnt, sSave, 255, 0, vbNullString, ByVal 0&, ByVal 0&) <> 0 Then Exit Do '退出循环
        '将枚举出来的子键逐一加入数组
        ReDim Preserve vName(0 To dNum, 0 To Cnt)
        vName(Num2, Cnt) = StripTerminator(sSave)
        Cnt = Cnt + 1
    Loop
End If

    '关闭这个注册表键
    RegCloseKey hKey
End Sub
'Public Sub GetAllValue(ByVal hResult As Long, ByVal strKey As String, ByRef vName As Variant, Optional ByVal dNum As Integer = -1, Optional ByVal Num2 As Integer = 0)
Public Sub GetAllValue(ByVal hResult As Long, ByVal strKey As String, ByRef vName As Collection, Optional ByVal dNum As Integer = -1, Optional ByVal Num2 As Integer = 0)
Dim lngFilterIndex As Long

    Dim hKey As Long, Cnt As Long, sSave As String
    '打开一个注册表键
    RegOpenKey hResult, strKey, hKey
    '列出该键下的所有子键
    Cnt = 0

If dNum < 0 Then '一维数组
    Do
        '创建一个缓冲区
        sSave = String(255, 0)
        '枚举出所有子键
        If RegEnumValue(hKey, Cnt, sSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        '将枚举出来的子键逐一加入数组
        'ReDim Preserve vName(0 To Cnt)
        'vName(Cnt) = StripTerminator(sSave)
        vName.Add StripTerminator(sSave)
        Cnt = Cnt + 1
    Loop
Else
    Dim vNameCild As Collection
    Do
        '创建一个缓冲区
        sSave = String(255, 0)
        '枚举出所有子键
        If RegEnumValue(hKey, Cnt, sSave, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do '退出循环
        '将枚举出来的子键逐一加入数组
        'ReDim Preserve vName(0 To dNum, 0 To Cnt)
        Set vNameCild = New Collection
        vNameCild.Add StripTerminator(sSave)
        vNameCild.Add GetString(hResult, strKey, StripTerminator(sSave))
        
        'vName(Num2, Cnt) = StripTerminator(sSave)
        'vName(Num2 + 1, Cnt) = GetString(hResult, strKey, StripTerminator(sSave))
        vName.Add vNameCild
        Cnt = Cnt + 1
    Loop
End If

    '关闭这个注册表键
    RegCloseKey hKey
End Sub

