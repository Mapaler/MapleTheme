Attribute VB_Name = "显示PNG"
Option Explicit
'*************************************************************************
'**模 块 名：ModPaintPNG
'**说    明：显示PNG图片的模块
'**创 建 人：嗷嗷叫的老马
'**日    期：2008年11月13日
'**版    本：V1.0
'**备    注：利用GDI显示PNG图片.PNG本身可实现半透明,比较省资源.
'**          紫水晶工作室 版权所有
'**          更多模块/类模块请访问我站:  http://www.m5home.com
'*************************************************************************

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Enum GpStatus
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single) As GpStatus
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, Image As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus

Dim gdip_Token&, gdip_pngImage&, gdip_Graphics&, Picname$

Public Sub PaintPng(ByVal sFileName As String, ByVal hDC As Long, ByVal mX As Long, ByVal mY As Long)
    '显示PNG图片到指定的DC环境
    '
    'mX与mY单位为象素.
    Dim lngHeight As Long, lngWidth As Long
   
    Call GDI_Initialize
    
    If GdipCreateFromHDC(hDC, gdip_Graphics) <> ok Then
        GdiplusShutdown gdip_Token
    Else
        Call GdipLoadImageFromFile(StrConv(GetShortName(sFileName), vbUnicode), gdip_pngImage)
        Call GdipGetImageHeight(gdip_pngImage, lngHeight)   '
        Call GdipGetImageWidth(gdip_pngImage, lngWidth)
        Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mX, mY, lngWidth, lngHeight)
    End If
    
    Call GDI_Terminate
End Sub

Public Sub PaintPng2(ByVal sFileName As String, ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0)
    '显示PNG图片到指定的DC环境
    '
    'mX与mY单位为象素.
    Dim lngHeight As Long, lngWidth As Long
    Dim showHeight As Single, showWidth As Single
    Dim Width_Scale As Single, Height_Scale As Single
    Dim mX As Long, mY As Long
    Call GDI_Initialize
    
    If GdipCreateFromHDC(hDC, gdip_Graphics) <> ok Then
        GdiplusShutdown gdip_Token
    Else
        Call GdipLoadImageFromFile(StrConv(GetShortName(sFileName), vbUnicode), gdip_pngImage)
        Call GdipGetImageHeight(gdip_pngImage, lngHeight)   '获得原图长宽
        Call GdipGetImageWidth(gdip_pngImage, lngWidth)
        If lngWidth = 0 Then lngWidth = 60 '部分无法正确读取的
        If lngHeight = 0 Then lngHeight = 45 '部分无法正确读取的
        Width_Scale = Width / lngWidth
        Height_Scale = Height / lngHeight
        If Width_Scale < Height_Scale Then
            showWidth = Width
            showHeight = lngHeight * Width_Scale
            mX = 0
            mY = (Height - showHeight) / 2 - 2
        Else
            showWidth = lngWidth * Height_Scale
            showHeight = Height
            mX = (Width - showWidth) / 2 - 2
            mY = 0
        End If
        Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mX + Left, mY + Top, showWidth, showHeight)
    End If
    
    Call GDI_Terminate
End Sub

Private Sub GDI_Initialize()
    Dim GpInput As GdiplusStartupInput
    
    GpInput.GdiplusVersion = 1
    gdip_Graphics = 0
    gdip_pngImage = 0
    If GdiplusStartup(gdip_Token, GpInput) <> ok Then
        Debug.Print "GDI初始失败！"
'        MsgBox "GDI初始失败！"
    End If
End Sub

Private Sub GDI_Terminate()
    GdipDisposeImage gdip_pngImage
    GdipDeleteGraphics gdip_Graphics
    GdiplusShutdown gdip_Token
End Sub

Private Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal&, sShortPathName$
    sShortPathName = Space(255)
    Call GetShortPathName(sLongFileName, sShortPathName, 255)
    If InStr(sShortPathName, Chr(0)) > 0 Then
        GetShortName = Trim(Mid(sShortPathName, 1, InStr(sShortPathName, Chr(0)) - 1))
    Else
        GetShortName = Trim(sShortPathName)
    End If
End Function


