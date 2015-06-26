Attribute VB_Name = "选择文件夹"
'VB也可以使用CallBack，下面是一个例子：
'先把下面的代码放入BAS模块：
Option Explicit

'common to both methods
Public Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib _
"shell32.dll" Alias "SHBrowseForFolderA" _
(lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib _
"shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, _
ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Public Declare Sub MoveMemory Lib "kernel32" _
Alias "RtlMoveMemory" _
(pDest As Any, _
pSource As Any, _
ByVal dwLength As Long)

Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SelectIONCHANGED
'message.
Public Const BFFM_SETSelectIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSelectIONW As Long = (WM_USER + 103)


'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Public Declare Function SHSimpleIDListFromPath Lib _
"shell32" Alias "#162" _
(ByVal szPath As String) As Long


'specific to the STRING method
Public Declare Function LocalAlloc Lib "kernel32" _
(ByVal uFlags As Long, _
ByVal uBytes As Long) As Long

Public Declare Function LocalFree Lib "kernel32" _
(ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
(lpString1 As Any, lpString2 As Any) As Long

Public Declare Function lstrlenA Lib "kernel32" _
(lpString As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)


Public Function BrowseCallbackProcStr(ByVal hwnd As Long, _
ByVal uMsg As Long, _
ByVal lParam As Long, _
ByVal lpData As Long) As Long

'Callback for the Browse STRING method.

'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.

Select Case uMsg
Case BFFM_INITIALIZED

Call SendMessage(hwnd, BFFM_SETSelectIONA, _
True, ByVal StrFromPtrA(lpData))

Case Else:

End Select

End Function


Public Function BrowseCallbackProc(ByVal hwnd As Long, _
ByVal uMsg As Long, _
ByVal lParam As Long, _
ByVal lpData As Long) As Long

'Callback for the Browse PIDL method.

'On initialization, set the dialog's
'pre-selected folder using the pidl
'set as the bi.lParam, and passed back
'to the callback as lpData param.

Select Case uMsg
Case BFFM_INITIALIZED

Call SendMessage(hwnd, BFFM_SETSelectIONA, _
False, ByVal lpData)

Case Else:

End Select

End Function


Public Function FARPROC(pfn As Long) As Long

'A dummy procedure that receives and returns
'the value of the AddressOf operator.

'Obtain and set the address of the callback
'This workaround is needed as you can't assign
'AddressOf directly to a member of a user-
'defined type, but you can assign it to another
'long and use that (as returned here)

FARPROC = pfn

End Function


Public Function StrFromPtrA(lpszA As Long) As String

'Returns an ANSI string from a pointer to an ANSI string.

Dim sRtn As String
sRtn = String$(lstrlenA(ByVal lpszA), 0)
Call lstrcpyA(ByVal sRtn, ByVal lpszA)
StrFromPtrA = sRtn

End Function

'--end block--'
'按PIDL
Public Function BrowseForFolderByPIDL(sSelPath As String) As String
Dim BI As BROWSEINFO
Dim pidl As Long
Dim sPath As String * MAX_PATH

With BI
.hOwner = Main.hwnd 'Me.hWnd
.pidlRoot = 0
.lpszTitle = "Pre-selecting a folder using the folder's pidl."
.lpfn = FARPROC(AddressOf BrowseCallbackProc)
.lParam = SHSimpleIDListFromPath(sSelPath)
End With

pidl = SHBrowseForFolder(BI)

If pidl Then
If SHGetPathFromIDList(pidl, sPath) Then
BrowseForFolderByPIDL = Left$(sPath, InStr(sPath, vbNullChar) - 1)
End If
Call CoTaskMemFree(pidl)
End If

Call CoTaskMemFree(BI.lParam)
End Function

'按路径打开文件夹
Public Function BrowseForFolderByPath(sSelPath As String, ByVal ShowString As String, ByVal Form_Me As Variant) As String
Dim BI As BROWSEINFO
Dim pidl As Long
Dim lpSelPath As Long
Dim sPath As String * MAX_PATH

With BI
.hOwner = Form_Me.hwnd 'Me.hWnd
.pidlRoot = 0
.lpszTitle = ShowString
.lpfn = FARPROC(AddressOf BrowseCallbackProcStr)

lpSelPath = LocalAlloc(LPTR, Len(sSelPath))
MoveMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath)
.lParam = lpSelPath

End With

pidl = SHBrowseForFolder(BI)

If pidl Then
If SHGetPathFromIDList(pidl, sPath) Then
BrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
End If
Call CoTaskMemFree(pidl)
End If

Call LocalFree(lpSelPath)
End Function

'获得文件列表
Public Sub GetFileName(ByVal sPath As String, ByVal Filter As String, ByRef vName As Collection)  '这是获取指定文件夹下指定后缀名的文件名称的过程
Dim sDir As String
Dim sFilter() As String
Dim lngFilterIndex As Long
Dim lngIndex As Long
Dim FileExtension As String
'Dim lngfiles As Integer

Set vName = New Collection '先清空集合
'Erase vName '先清空数组

sFilter = Split(Filter, ",")
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

sDir = Dir(sPath & "*.*", vbNormal)

Do While Len(sDir) > 0
    FileExtension = Mid$(sDir, InStrRev(sDir, ".") + 1) '获取扩展名
    For lngFilterIndex = LBound(sFilter) To UBound(sFilter) '按需要的扩展名个数循环
        If FileExtension = sFilter(lngFilterIndex) Then '当扩展名是需要的扩展名的时候
            'lngfiles = lngfiles + 1
            'ReDim Preserve vName(1 To lngfiles)
            vName.Add sPath & sDir
            'vName(lngfiles) = sPath & sDir
            Exit For '退出本轮扩展名判断循环
        End If
    Next
    sDir = Dir
Loop
End Sub

'获得文件列表-递归
Public Sub GetFileNameRecursion(ByVal sPath As String, ByVal Filter As String, ByRef vName As Collection)   '这是获取指定文件夹下指定后缀名的文件名称的过程
Dim sDir As String, sDir2 As String
Dim sFilter() As String
Dim lngFilterIndex As Long
Dim lngIndex As Long
Dim FileExtension As String
'Dim lngfiles As Integer
sFilter = Split(Filter, ",")
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

'lngfiles = vName.count - 1
'lngfiles = UBound(vName)

sDir = Dir(sPath & "*.*", vbNormal + vbHidden + vbSystem)
Do While Len(sDir) > 0
    FileExtension = Mid$(sDir, InStrRev(sDir, ".") + 1) '获取扩展名
    For lngFilterIndex = LBound(sFilter) To UBound(sFilter) '按需要的扩展名个数循环
        If FileExtension = sFilter(lngFilterIndex) Then '当扩展名是需要的扩展名的时候
            'lngfiles = lngfiles + 1
            'ReDim Preserve vName(0 To lngfiles)
            vName.Add sPath & sDir
            'vName(lngfiles) = sPath & sDir
            Exit For '退出本轮扩展名判断循环
        End If
    Next
    sDir = Dir
Loop

'开始递归文件夹
Dim ChildrenFolder As New Collection, FIndex As Long
'FIndex = 0
'ReDim Preserve ChildrenFolder(0)
sDir2 = Dir(sPath, vbDirectory + vbHidden + vbSystem)
Do While Len(sDir2) > 0
    If sDir2 <> "." And sDir2 <> ".." Then
        If (GetAttr(sPath & sDir2) And vbDirectory) = vbDirectory Then
           ' FIndex = FIndex + 1
            'ReDim Preserve ChildrenFolder(0 To FIndex)
            ChildrenFolder.Add sDir2
            'ChildrenFolder(FIndex) = sDir2
        End If
    End If
    sDir2 = Dir
Loop
For FIndex = 1 To ChildrenFolder.count
    GetFileNameRecursion sPath & ChildrenFolder(FIndex), Filter, vName
Next
End Sub
