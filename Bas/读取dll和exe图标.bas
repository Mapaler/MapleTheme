Attribute VB_Name = "读取dll和exe图标"
Option Explicit

Global lIcon As Long
'Download by http://www.codefans.net
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hicon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hicon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'读取dll文本内容
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'本程序结合了下面两人的例子，并自己加工了点
'A: KPD-Team 1999, 2001 URL: http://www.allapi.net/ E-Mail: KPDTeam@Allapi.net additional coding by Willem Bogaerts, w-p@dds.nl
'B: Example by Shafru (shafru@hotmail.com)
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHChangeIconDialog Lib "shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long
'Detect if the program is running under Windows NT
Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
Public Function chooseIcon(ByRef strFile As String, ByRef lngIconNum As Long, Form_Me As Form) As Boolean
'这个Form_Me As Form是为了让这个程序能放在模块里好调用
    Dim str1 As String * 260
    Dim lng1 As Long ' Dummy?
    Dim lngResult As Long
    
    str1 = strFile & vbNullChar
    '判断系统是否是NT内核
    If IsWinNT Then
        'if we're in WinNT, we have to call the Unicode version of the function
        str1 = StrConv(str1, vbUnicode)
        lngResult = SHChangeIconDialog(Form_Me.hwnd, str1, Len(str1), lngIconNum)
        str1 = StrConv(str1, vbFromUnicode)
    Else
        'if we're in Win9x, we have to call the ANSI version of the function
        lngResult = SHChangeIconDialog(Form_Me.hwnd, str1, lng1, lngIconNum)
    End If
    
    '这个程序自己的返回值 0 (失败) 或 1 (成功)
    '将str1改编为选择的文件名
    chooseIcon = (lngResult <> 0)
    If chooseIcon Then
        strFile = Left(str1, InStr(str1, vbNullChar) - 1)
    End If
End Function


'获取dll文本内容
Public Function Get_dll_text(ByVal text_temp As String) As String
If InStr(text_temp, "@") > 0 And InStr(text_temp, ",") > 0 Then
    text_temp = url_to_N(Right(text_temp, Len(text_temp) - 1)) '去掉@号再变为普通路径
    
    Dim buffer As String * 255
    Dim Temp As Long
    Dim dll_text_num As Long
    Dim url_temp As String
    
        url_temp = Left(text_temp, InStr(text_temp, ",") - 1)
        dll_text_num = Mid(text_temp, InStr(text_temp, ",") + 2) '后面的-1是为了去掉负号
        Temp = LoadString(LoadLibrary(url_temp), dll_text_num, buffer, 255)
        Get_dll_text = StripTerminator(buffer)
Else
    Get_dll_text = text_temp
End If
End Function
