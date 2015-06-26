Attribute VB_Name = "全玻璃窗口"
'****************************************************************************************
'*************  请在Windows Vista 下调试/运行本程序，否则会看到古怪的效果 ***************
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

'获取系统颜色
Public Const COLOR_SCROLLBAR = 0 '滚动条
Public Const COLOR_BACKGROUND = 1 '桌面背景
Public Const COLOR_ACTIVECAPTION = 2 '活动窗口标题
Public Const COLOR_INACTIVECAPTION = 3 '非活动窗口标题
Public Const COLOR_MENU = 4 '菜单
Public Const COLOR_WINDOW = 5 '窗口背景
Public Const COLOR_WINDOWFRAME = 6 '窗口框
Public Const COLOR_MENUTEXT = 7 '窗口文字
Public Const COLOR_WINDOWTEXT = 8 '3D阴影
Public Const COLOR_CAPTIONTEXT = 9 '标题文字
Public Const COLOR_ACTIVEBORDER = 10 '活动窗口边框
Public Const COLOR_INACTIVEBORDER = 11 '非活动窗口边框
Public Const COLOR_APPWORKSPACE = 12 'MDI窗口背景
Public Const COLOR_HIGHLIGHT = 13 '选择条背景
Public Const COLOR_HIGHLIGHTTEXT = 14 '选择条文字
Public Const COLOR_BTNFACE = 15 '按钮
Public Const COLOR_BTNSHADOW = 16 '3D按钮阴影
Public Const COLOR_GRAYTEXT = 17 '灰度文字
Public Const COLOR_BTNTEXT = 18 '按钮文字
Public Const COLOR_INACTIVECAPTIONTEXT = 19 '非活动窗口文字
Public Const COLOR_BTNHIGHLIGHT = 20 '3D选择按钮
Public Declare Function SetSysColors Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

