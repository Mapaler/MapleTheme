Attribute VB_Name = "ȫ��������"
'****************************************************************************************
'*************  ����Windows Vista �µ���/���б����򣬷���ῴ���Źֵ�Ч�� ***************
'****************************************************************************************


Public Type MARGINS
 'public int Left;
m_Left As Long
  'public int Right;
m_Right As Long
  'public int Top;
m_Top As Long
 'public int Bottom;
m_Button As Long


End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Dim Inied As Boolean
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern void DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS margins);
Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, margin As MARGINS) As Long
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern bool DwmIsCompositionEnabled();
Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetLayeredWindowAttributesByColor Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

'��ȡϵͳ��ɫ
Public Const COLOR_SCROLLBAR = 0 '������
Public Const COLOR_BACKGROUND = 1 '���汳��
Public Const COLOR_ACTIVECAPTION = 2 '����ڱ���
Public Const COLOR_INACTIVECAPTION = 3 '�ǻ���ڱ���
Public Const COLOR_MENU = 4 '�˵�
Public Const COLOR_WINDOW = 5 '���ڱ���
Public Const COLOR_WINDOWFRAME = 6 '���ڿ�
Public Const COLOR_MENUTEXT = 7 '��������
Public Const COLOR_WINDOWTEXT = 8 '3D��Ӱ
Public Const COLOR_CAPTIONTEXT = 9 '��������
Public Const COLOR_ACTIVEBORDER = 10 '����ڱ߿�
Public Const COLOR_INACTIVEBORDER = 11 '�ǻ���ڱ߿�
Public Const COLOR_APPWORKSPACE = 12 'MDI���ڱ���
Public Const COLOR_HIGHLIGHT = 13 'ѡ��������
Public Const COLOR_HIGHLIGHTTEXT = 14 'ѡ��������
Public Const COLOR_BTNFACE = 15 '��ť
Public Const COLOR_BTNSHADOW = 16 '3D��ť��Ӱ
Public Const COLOR_GRAYTEXT = 17 '�Ҷ�����
Public Const COLOR_BTNTEXT = 18 '��ť����
Public Const COLOR_INACTIVECAPTIONTEXT = 19 '�ǻ��������
Public Const COLOR_BTNHIGHLIGHT = 20 '3Dѡ��ť
Public Declare Function SetSysColors Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

