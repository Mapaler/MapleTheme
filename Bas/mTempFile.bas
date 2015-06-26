Attribute VB_Name = "mTempFile"
Option Explicit

' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
'  dwPlatformId defines:
'Private Const VER_PLATFORM_WIN32s = 0
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'Private Const VER_PLATFORM_WIN32_NT = 2


Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_PATH = 260

Private Type OSVERSIONINFOEX
    dwOSVersionExInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const PRODUCT_UNLICENSED As Long = &HABCDABCD
Private Const PRODUCT_BUSINESS As Long = &H6
Private Const PRODUCT_BUSINESS_N As Long = &H10
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC
Private Const PRODUCT_ENTERPRISE As Long = &H4
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF
Private Const PRODUCT_HOME_BASIC As Long = &H2
Private Const PRODUCT_HOME_BASIC_N As Long = &H5
Private Const PRODUCT_HOME_PREMIUM As Long = &H3
Private Const PRODUCT_HOME_PREMIUM_N As Long = &H1A
Private Const PRODUCT_HOME_SERVER As Long = &H13
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Long = &H18
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19
Private Const PRODUCT_STANDARD_SERVER As Long = &H7
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD
Private Const PRODUCT_STARTER As Long = &H8
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Long = &H17
Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Long = &H14
Private Const PRODUCT_STORAGE_STANDARD_SERVER As Long = &H15
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Long = &H16
Private Const PRODUCT_UNDEFINED As Long = &H0
Private Const PRODUCT_ULTIMATE As Long = &H1
Private Const PRODUCT_ULTIMATE_N As Long = &H1C
Private Const PRODUCT_WEB_SERVER As Long = &H11


Public Sub Get_Ststem_Ver() '读取系统版本
'Dim strComputer, objWMIService, colItems, objItem
'    strComputer = "."
'
'    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
'
'    For Each objItem In colItems
'        strOSversion = objItem.Version
'    Next
'
'    System_Ver = Left(strOSversion, 3)
'
'    Select Case Left(strOSversion, 3)
'    Case Is < 5
'        strOSversion = "Windows ME or Old"
'    Case "5.0"
'        strOSversion = "Windows 2000"
'    Case "5.1"
'        strOSversion = "Windows XP"
'    Case "5.2"
'        strOSversion = "Windows Server 2003"
'    Case "6.0"
'        strOSversion = "Windows visita"
'    Case "6.1"
'        strOSversion = "Windows 7"
'    Case Is >= 6.2
'        strOSversion = "Windows 8 or New"
'    Case Else
'        strOSversion = "Don't know"
'    End Select
    strOSversion = GetOSVersionEx()
    ' 变量声明
    Dim retLng As Long, OSVersionEx As OSVERSIONINFOEX
    '结构尺寸
    OSVersionEx.dwOSVersionExInfoSize = Len(OSVersionEx)
    '获取 Windows 版本
    retLng = GetVersionEx(OSVersionEx)
    If retLng = 0 Then
        System_Ver = 0
        Exit Sub
    End If
    With OSVersionEx
        System_Ver = .dwMajorVersion + .dwMinorVersion / 10
    End With
End Sub

Public Function GetOSVersionEx() As String
    ' 变量声明
    Dim retLng As Long, OSVersionEx As OSVERSIONINFOEX
    '结构尺寸
    OSVersionEx.dwOSVersionExInfoSize = Len(OSVersionEx)
    '获取 Windows 版本
    retLng = GetVersionEx(OSVersionEx)
    If retLng = 0 Then
        GetOSVersionEx = "未知"
        Exit Function
    End If
    With OSVersionEx
        Select Case .dwPlatformId
        Case VER_PLATFORM_WIN32s
            Select Case .dwMajorVersion
            Case 1
                Select Case .dwMinorVersion
                Case 0 ' Win 1.0
                    GetOSVersionEx = "Windows 1.0"
                Case Else
                    GetOSVersionEx = "Win32s"
                End Select
            Case 2
                Select Case .dwMinorVersion
                Case 0 ' Win 2.0
                    GetOSVersionEx = "Windows 2.0"
                Case Else
                    GetOSVersionEx = "Win32s"
                End Select
            Case 3
                Select Case .dwMinorVersion
                Case 0 ' Win 3.0
                    GetOSVersionEx = "Windows 3.0"
                Case 1 ' Win 3.1
                    GetOSVersionEx = "Windows 3.1"
                Case 2 ' Win 3.2
                    GetOSVersionEx = "Windows 3.2"
                Case Else
                    GetOSVersionEx = "Windows 3.x"
                End Select
            Case Else
                GetOSVersionEx = "Win32s"
            End Select
        Case VER_PLATFORM_WIN32_WINDOWS
            Select Case .dwMajorVersion
            Case 4
                Select Case .dwMinorVersion
                Case 0 ' Win 95
                    Select Case .szCSDVersion
                    Case "C" ' OSR2
                        GetOSVersionEx = "Windows 95 OSR2"
                    Case "B" ' OSR2
                        GetOSVersionEx = "Windows 95 OSR2"
                    Case Else
                        GetOSVersionEx = "Windows 95"
                    End Select
                Case 1998 ' Win 98
                    GetOSVersionEx = "Windows 98"
                Case 10 ' Win 98
                    Select Case .szCSDVersion
                    Case "A" ' SE
                        GetOSVersionEx = "Windows 98 SE"
                    Case "C" ' OSR2
                        GetOSVersionEx = "Windows 98 OSR2"
                    Case "B" ' OSR2
                        GetOSVersionEx = "Windows 98 OSR2"
                    Case Else
                        GetOSVersionEx = "Windows 98"
                    End Select
                Case 90 ' Win ME
                    GetOSVersionEx = "Windows ME"
                End Select
            Case Else
                GetOSVersionEx = "Win32"
            End Select
        Case VER_PLATFORM_WIN32_NT
            Select Case .dwMajorVersion
            Case 5
                Select Case .dwMinorVersion
                Case 0 ' Win 2000
                    Select Case .wProductType
                    Case 1
                        Select Case .wSuiteMask
                        Case &H80 ' 数据中心
                            GetOSVersionEx = "Windows 2000 Data center"
                        Case &H2 ' 高级版本
                            GetOSVersionEx = "Windows 2000 Advanced"
                        Case Else
                            GetOSVersionEx = "Windows 2000"
                        End Select
                    End Select
                Case 1 ' Win XP
                    Select Case .wProductType
                    Case 1
                        Select Case .wSuiteMask
                        Case &H0 ' Pro
                            GetOSVersionEx = "Windows XP Professional"
                        Case &H200 ' Home
                            GetOSVersionEx = "Windows XP Home"
                        Case Else ' XP
                            GetOSVersionEx = "Windows XP"
                        End Select
                    End Select
                Case 2 ' Win Server 2003
                    Select Case .wProductType
                    Case 3
                        Select Case .wSuiteMask
                        Case &H2
                            GetOSVersionEx = "Windows Server 2003 Enterprise"
                        Case &H80
                            GetOSVersionEx = "Windows Server 2003 Data center"
                        Case &H400
                            GetOSVersionEx = "Windows Server 2003 Web Edition"
                        Case &H0
                            GetOSVersionEx = "Windows Server 2003 Standard"
                        Case Else
                            GetOSVersionEx = "Windows Server 2003"
                        End Select
                    End Select
                Case Else
                    GetOSVersionEx = "Windows NT"
                End Select
            Case 6
                Select Case .wProductType
                Case PRODUCT_BUSINESS
                    GetOSVersionEx = "Business Edition"
                Case PRODUCT_BUSINESS_N
                    GetOSVersionEx = "Business Edition (N)"
                Case PRODUCT_CLUSTER_SERVER
                    GetOSVersionEx = "Cluster Server Edition"
                Case PRODUCT_DATACENTER_SERVER
                    GetOSVersionEx = "Server Datacenter Edition (Full Installation)"
                Case PRODUCT_DATACENTER_SERVER_CORE
                    GetOSVersionEx = "Server Datacenter Edition (Core Installation)"
                Case PRODUCT_ENTERPRISE
                    GetOSVersionEx = "Enterprise Edition"
                Case PRODUCT_ENTERPRISE_N
                    GetOSVersionEx = "Enterprise Edition (N)"
                Case PRODUCT_ENTERPRISE_SERVER
                    GetOSVersionEx = "Server Enterprise Edition (Full Installation)"
                Case PRODUCT_ENTERPRISE_SERVER_CORE
                    GetOSVersionEx = "Server Enterprise Edition (Core Installation)"
                Case PRODUCT_ENTERPRISE_SERVER_IA64
                    GetOSVersionEx = "Server Enterprise Edition for Itanium-based Systems"
                Case PRODUCT_HOME_BASIC
                    GetOSVersionEx = "Home Basic Edition"
                Case PRODUCT_HOME_BASIC_N
                    GetOSVersionEx = "Home Basic Edition (N)"
                Case PRODUCT_HOME_PREMIUM
                    GetOSVersionEx = "Home Premium Edition"
                Case PRODUCT_HOME_PREMIUM_N
                    GetOSVersionEx = "Home Premium Edition (N)"
                Case PRODUCT_HOME_SERVER
                    GetOSVersionEx = "Home Server Edition"
                Case PRODUCT_SERVER_FOR_SMALLBUSINESS
                    GetOSVersionEx = "Server for Small Business Edition"
                Case PRODUCT_SMALLBUSINESS_SERVER
                    GetOSVersionEx = "Small Business Server"
                Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM
                    GetOSVersionEx = "Small Business Server Premium Edition"
                Case PRODUCT_STANDARD_SERVER
                    GetOSVersionEx = "Server Standard Edition (Full Installation)"
                Case PRODUCT_STANDARD_SERVER_CORE
                    GetOSVersionEx = "Server Standard Edition (Core Installation)"
                Case PRODUCT_STARTER
                    GetOSVersionEx = "Starter Edition"
                Case PRODUCT_STORAGE_ENTERPRISE_SERVER
                    GetOSVersionEx = "Storage Server Enterprise Edition"
                Case PRODUCT_STORAGE_EXPRESS_SERVER
                    GetOSVersionEx = "Storage Server Express Edition"
                Case PRODUCT_STORAGE_STANDARD_SERVER
                    GetOSVersionEx = "Storage Server Standard Edition"
                Case PRODUCT_STORAGE_WORKGROUP_SERVER
                    GetOSVersionEx = "Storage Server Workgroup Edition"
                Case PRODUCT_ULTIMATE
                    GetOSVersionEx = "Ultimate Edition"
                Case PRODUCT_ULTIMATE_N
                    GetOSVersionEx = "Ultimate Edition (N)"
                Case PRODUCT_UNDEFINED
                    GetOSVersionEx = "Unknown product"
                Case PRODUCT_UNLICENSED
                    GetOSVersionEx = "Not activated product"
                Case PRODUCT_WEB_SERVER
                    GetOSVersionEx = "Web Server Edition"
                End Select
                Select Case .dwMinorVersion
                Case 0
                    Select Case .wProductType
                    Case 1 ' Win Vista
                        GetOSVersionEx = "Windows Vista " & GetOSVersionEx
                    Case 3 ' Win Server 2008
                        GetOSVersionEx = "Windows Server 2008 " & GetOSVersionEx
                    Case Else
                        GetOSVersionEx = "Windows Vista " & GetOSVersionEx
                    End Select
                Case 1
                    Select Case .wProductType
                    Case 1 ' Win 7
                        GetOSVersionEx = "Windows 7 " & GetOSVersionEx
                    Case 3 ' Win Server 2008 R2
                        GetOSVersionEx = "Windows Server 2008 R2 " & GetOSVersionEx
                    Case Else
                        GetOSVersionEx = "Windows 7 " & GetOSVersionEx
                    End Select
                Case 2
                    Select Case .wProductType
                    Case 1 ' Win 8
                        GetOSVersionEx = "Windows 8 " & GetOSVersionEx
                    Case 3 ' Win Server 2012
                        GetOSVersionEx = "Windows Server 2012 " & GetOSVersionEx
                    Case Else
                        GetOSVersionEx = "Windows 8 " & GetOSVersionEx
                    End Select
                Case Is >= 3
                    Select Case .wProductType
                    Case 1 ' Win 8.1
                        GetOSVersionEx = "Windows 8.1 " & GetOSVersionEx
                    Case 3 ' Win Server 2012 R2
                        GetOSVersionEx = "Windows Server 2012 R2 " & GetOSVersionEx
                    Case Else
                        GetOSVersionEx = "Windows 8.1 " & GetOSVersionEx
                    End Select
                Case Else
                    GetOSVersionEx = "Windows NT"
                End Select 'Minor
            Case Is > 6
                GetOSVersionEx = "Windows 10"
            Case Else
                GetOSVersionEx = "Windows NT"
            End Select 'Major
        Case Else
            GetOSVersionEx = "Unknown Platform"
        End Select 'Platform
        If .wServicePackMajor > 0 Then
            GetOSVersionEx = GetOSVersionEx & " Service Pack " & .wServicePackMajor & IIf(.wServicePackMinor > 0, "." & .wServicePackMinor, vbNullString)
        End If
        GetOSVersionEx = GetOSVersionEx & " [Version:" & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & "]"
    End With
End Function
Public Property Get TempDir() As String
Dim sRet As String, c As Long
Dim lErr As Long
   sRet = String$(MAX_PATH, 0)
   c = GetTempPath(MAX_PATH, sRet)
   lErr = Err.LastDllError
   If c = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   TempDir = Left$(sRet, c)
End Property
Public Property Get TempFileName( _
        Optional ByVal sPrefix As String, _
        Optional ByVal sPathName As String) As String
Dim lErr As Long
Dim iPos As Long

   If sPrefix = "" Then sPrefix = ""
   If sPathName = "" Then sPathName = TempDir
   
   Dim sRet As String
   sRet = String(MAX_PATH, 0)
   GetTempFileName sPathName, sPrefix, 0, sRet
   lErr = Err.LastDllError
   If Not lErr = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   iPos = InStr(sRet, vbNullChar)
   If Not iPos = 0 Then
      TempFileName = Left$(sRet, iPos - 1)
   End If
End Property


Public Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
   
   ' Return the error message associated with LastDLLError:
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If
   
End Function


Public Function IsNT() As Boolean
Dim tV As OSVERSIONINFOEX
   GetVersionEx tV
   IsNT = (tV.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function


