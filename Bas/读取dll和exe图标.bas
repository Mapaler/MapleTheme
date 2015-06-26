Attribute VB_Name = "��ȡdll��exeͼ��"
Option Explicit

Global lIcon As Long
'Download by http://www.codefans.net
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hicon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hicon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'��ȡdll�ı�����
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'�����������������˵����ӣ����Լ��ӹ��˵�
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
'���Form_Me As Form��Ϊ������������ܷ���ģ����õ���
    Dim str1 As String * 260
    Dim lng1 As Long ' Dummy?
    Dim lngResult As Long
    
    str1 = strFile & vbNullChar
    '�ж�ϵͳ�Ƿ���NT�ں�
    If IsWinNT Then
        'if we're in WinNT, we have to call the Unicode version of the function
        str1 = StrConv(str1, vbUnicode)
        lngResult = SHChangeIconDialog(Form_Me.hwnd, str1, Len(str1), lngIconNum)
        str1 = StrConv(str1, vbFromUnicode)
    Else
        'if we're in Win9x, we have to call the ANSI version of the function
        lngResult = SHChangeIconDialog(Form_Me.hwnd, str1, lng1, lngIconNum)
    End If
    
    '��������Լ��ķ���ֵ 0 (ʧ��) �� 1 (�ɹ�)
    '��str1�ı�Ϊѡ����ļ���
    chooseIcon = (lngResult <> 0)
    If chooseIcon Then
        strFile = Left(str1, InStr(str1, vbNullChar) - 1)
    End If
End Function


'��ȡdll�ı�����
Public Function Get_dll_text(ByVal text_temp As String) As String
If InStr(text_temp, "@") > 0 And InStr(text_temp, ",") > 0 Then
    text_temp = url_to_N(Right(text_temp, Len(text_temp) - 1)) 'ȥ��@���ٱ�Ϊ��ͨ·��
    
    Dim buffer As String * 255
    Dim Temp As Long
    Dim dll_text_num As Long
    Dim url_temp As String
    
        url_temp = Left(text_temp, InStr(text_temp, ",") - 1)
        dll_text_num = Mid(text_temp, InStr(text_temp, ",") + 2) '�����-1��Ϊ��ȥ������
        Temp = LoadString(LoadLibrary(url_temp), dll_text_num, buffer, 255)
        Get_dll_text = StripTerminator(buffer)
Else
    Get_dll_text = text_temp
End If
End Function
