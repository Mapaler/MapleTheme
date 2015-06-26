Attribute VB_Name = "启动预设值"
Option Explicit
Dim i%, j%, k%, n%
'保持窗体在前
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


Public Sub Get_Options() '读取设置
Dim Lanuage_ShortName As String
    'Erase Lanuages '先清空数组
    Call GetFileName(url_to_N(App.Path & "\Lanuages"), "ini", Lanuages) '获取语言文件列表
    Lanuage_Now = 0
If Dir(Config_Url) <> "" Then
'On Error GoTo ErrHandler
    Lanuage_ShortName = GetFromIni("Option", "Lanuage", Config_Url) '获取短名
    'For i = 1 To UBound(Lanuages)
    For i = 1 To Lanuages.count
        If Lanuage_ShortName = GetFromIni("info", "ShortName", Lanuages(i)) Then
            Lanuage_Now = i '设置当前语言
            Exit For
        End If
    Next i
'ErrHandler: '下标错误（没有语言文件）
If GetFromIni("Option", "SystemTextShow", Config_Url) <> "" Then
    SystemTextShow = GetFromIni("Option", "SystemTextShow", Config_Url)
Else
    SystemTextShow = 0
End If
If GetFromIni("Option", "SoundList", Config_Url) <> "" Then
    Sound_Style = GetFromIni("Option", "SoundList", Config_Url)
Else
    Sound_Style = 0
End If
If GetFromIni("Option", "Aero", Config_Url) <> "" Then
    glass_ok = GetFromIni("Option", "Aero", Config_Url)
Else
    glass_ok = False
End If
    
If GetFromIni("Option", "AutoPaper", Config_Url) <> "" Then
    AutoPaper = GetFromIni("Option", "AutoPaper", Config_Url)
Else
    AutoPaper = 2
End If
'如果不是Aero就不开Aero
If System_Ver < 6 Or GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive") = 0 Or GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "Composition") = 0 Then
    glass_ok = False
End If
If GetFromIni("Option", "Aplha_Back_Color", Config_Url) <> "" Then
    Aplha_Back_Color = x16_to_x10(GetFromIni("Option", "Aplha_Back_Color", Config_Url))
Else
    Aplha_Back_Color = x16_to_x10("FFFFFF")
End If
    
    '判断是否是真是的文件目录地址
    If Is_File_Directory(GetFromIni("Option", "SystemRoot", Config_Url)) Then
        SysRoot = GetFromIni("Option", "SystemRoot", Config_Url)
    Else
        SysRoot = 0
    End If
    
    If GetFromIni("Option", "SysPath_Default", Config_Url) <> "" Then
        SysPath = GetFromIni("Option", "SysPath_Default", Config_Url)
    Else
        SysPath = 0
    End If
    
Else
    MsgBox "没有检测到设置文件，请选择软件语言。" + vbLf + "No found config.ini.Please choose soft lanuage.", 64, "No found Options"
    Options.Show
    frmLoad.Hide
End If
End Sub

'建立各种默认值
Public Sub Creat_Default()
Dim Root As Node

BAT_Color(0) = &H0&
BAT_Color(1) = &H800000
BAT_Color(2) = &H8000&
BAT_Color(3) = &H808000
BAT_Color(4) = &H80&
BAT_Color(5) = &H800080
BAT_Color(6) = &H8080&
BAT_Color(7) = &HC0C0C0
BAT_Color(8) = &H808080
BAT_Color(9) = &HFF0000
BAT_Color(10) = &HFF00&
BAT_Color(11) = &HFFFF00
BAT_Color(12) = &HFF&
BAT_Color(13) = &HFF00FF
BAT_Color(14) = &HFFFF&
BAT_Color(15) = &HFFFFFF

'给图标列表数组添加数
'SysIco(0, 1) = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
'SysIco(1, 1) = "{59031A47-3F72-44A7-89C5-5595FE6B30EE}"
'SysIco(2, 1) = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
'SysIco(3, 1) = "{645FF040-5081-101B-9F08-00AA002F954E}"
'SysIco(4, 1) = "{645FF040-5081-101B-9F08-00AA002F954E}"
'SysIco(5, 1) = "{871C5380-42A0-1069-A2EA-08002B30309D}"

If SystemTextShow = True Then
    SysIco(0, 2) = Load_Lanuage("我的电脑", "Main", "Icon_Name0", Lanuage_Now)
    SysIco(1, 2) = Load_Lanuage("我的文档", "Main", "Icon_Name1", Lanuage_Now)
    SysIco(2, 2) = Load_Lanuage("网上邻居", "Main", "Icon_Name2", Lanuage_Now)
    SysIco(3, 2) = Load_Lanuage("回收站（空）", "Main", "Icon_Name3", Lanuage_Now)
    SysIco(4, 2) = Load_Lanuage("回收站（满）", "Main", "Icon_Name4", Lanuage_Now)
    SysIco(5, 2) = Load_Lanuage("Internet Explorer", "Main", "Icon_Name5", Lanuage_Now)
Else
SysIco(0, 2) = "@%SystemRoot%\system32\shell32.dll,-30386"
SysIco(1, 2) = "@%SystemRoot%\system32\shell32.dll,-30391"
SysIco(2, 2) = "@%SystemRoot%\system32\shell32.dll,-30387"
SysIco(3, 2) = "@%SystemRoot%\system32\shell32.dll,-30388"
SysIco(4, 2) = "@%SystemRoot%\system32\shell32.dll,-30389"
SysIco(5, 2) = "Internet Explorer"
End If

SysIco(5, 3) = "%ProgramFiles%\Internet Explorer\iexplore.exe,-0" 'IE
If System_Ver < 6 Then
    SysIco(0, 3) = "%SystemRoot%\System32\Shell32.dll,15" '我的电脑
    SysIco(1, 3) = "%SystemRoot%\SYSTEM32\mydocs.dll," '我的文档
    SysIco(2, 3) = "%SystemRoot%\System32\Shell32.dll,17" '网上邻居
    SysIco(3, 3) = "%SystemRoot%\System32\shell32.dll,31" '回收站（空）
    SysIco(4, 3) = "%SystemRoot%\System32\shell32.dll,32" '回收站（满）
Else
    SysIco(0, 3) = "%SystemRoot%\System32\imageres.dll,-109" '我的电脑
    SysIco(1, 3) = "%SystemRoot%\System32\imageres.dll,-123" '我的文档
    SysIco(2, 3) = "%SystemRoot%\System32\imageres.dll,-25" '网上邻居
    SysIco(3, 3) = "%SystemRoot%\System32\imageres.dll,-55" '回收站（空）
    SysIco(4, 3) = "%SystemRoot%\System32\imageres.dll,-54" '回收站（满）
End If
'给鼠标列表数组添加数
SysCur(0, 1) = "Arrow": SysCur(0, 2) = "%SystemRoot%\cursors\aero_arrow.cur"
SysCur(1, 1) = "Help": SysCur(1, 2) = "%SystemRoot%\cursors\aero_helpsel.cur"
SysCur(2, 1) = "AppStarting": SysCur(2, 2) = "%SystemRoot%\cursors\aero_working.ani"
SysCur(3, 1) = "Wait": SysCur(3, 2) = "%SystemRoot%\cursors\aero_busy.ani"
SysCur(4, 1) = "Crosshair": SysCur(4, 2) = ""
SysCur(5, 1) = "IBeam": SysCur(5, 2) = ""
SysCur(6, 1) = "NWPen": SysCur(6, 2) = "%SystemRoot%\cursors\aero_pen.cur"
SysCur(7, 1) = "No": SysCur(7, 2) = "%SystemRoot%\cursors\aero_unavail.cur"
SysCur(8, 1) = "SizeNS": SysCur(8, 2) = "%SystemRoot%\cursors\aero_ns.cur"
SysCur(9, 1) = "SizeWE": SysCur(9, 2) = "%SystemRoot%\cursors\aero_ew.cur"
SysCur(10, 1) = "SizeNWSE": SysCur(10, 2) = "%SystemRoot%\cursors\aero_nwse.cur"
SysCur(11, 1) = "SizeNESW": SysCur(11, 2) = "%SystemRoot%\cursors\aero_nesw.cur"
SysCur(12, 1) = "SizeAll": SysCur(12, 2) = "%SystemRoot%\cursors\aero_move.cur"
SysCur(13, 1) = "UpArrow": SysCur(13, 2) = "%SystemRoot%\cursors\aero_up.cur"
SysCur(14, 1) = "Hand": SysCur(14, 2) = "%SystemRoot%\cursors\aero_link.cur"
If SystemTextShow = True Then
    If Lanuage_Now = 0 Then
        SysCur(0, 3) = "正常选择"
        SysCur(1, 3) = "帮助选择"
        SysCur(2, 3) = "后台运行"
        SysCur(3, 3) = "忙"
        SysCur(4, 3) = "精确选择"
        SysCur(5, 3) = "文本选择"
        SysCur(6, 3) = "手写"
        SysCur(7, 3) = "不可用"
        SysCur(8, 3) = "垂直调整"
        SysCur(9, 3) = "水平调整"
        SysCur(10, 3) = "沿对角线调整1"
        SysCur(11, 3) = "沿对角线调整2"
        SysCur(12, 3) = "移动"
        SysCur(13, 3) = "候选"
        SysCur(14, 3) = "连接选择"
    Else
        For i = 0 To 14
            SysCur(i, 3) = Load_Lanuage(SysCur(i, 1), "Main", "Mouse_Name" & i, Lanuage_Now)
        Next
    End If
Else
    SysCur(0, 3) = "@main.cpl,-207"
    SysCur(1, 3) = "@main.cpl,-218"
    SysCur(2, 3) = "@main.cpl,-209"
    SysCur(3, 3) = "@main.cpl,-208"
    SysCur(4, 3) = "@main.cpl,-212"
    SysCur(5, 3) = "@main.cpl,-211"
    SysCur(6, 3) = "@main.cpl,-219"
    SysCur(7, 3) = "@main.cpl,-210"
    SysCur(8, 3) = "@main.cpl,-213"
    SysCur(9, 3) = "@main.cpl,-214"
    SysCur(10, 3) = "@main.cpl,-215"
    SysCur(11, 3) = "@main.cpl,-216"
    SysCur(12, 3) = "@main.cpl,-217"
    SysCur(13, 3) = "@main.cpl,-220"
    SysCur(14, 3) = "@main.cpl,-205"
End If
'给音效列表添加内容
'添加列表项，所有有值的节点key为sX
With Main.TreeView_Sound.Nodes
    .Clear '清除以前的节点

Dim S_Sound() As String '子节点

Call GetAllKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Names", Sound_Name)
    ReDim Preserve Sound_Name(0 To UBound(Sound_Name) + 1)
    For i = UBound(Sound_Name) To 1 Step -1
        Sound_Name(i) = Sound_Name(i - 1)
    Next
    Sound_Name(i) = ".Current"
    Main.Combo_Sys_Snd.Clear '先清除
    For i = 0 To UBound(Sound_Name) '获得声音预置个数
        Call Main.Combo_Sys_Snd.AddItem(Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(i), vbNullString)), i)
    Next
    Main.Combo_Sys_Snd.list(0) = Load_Lanuage("系统当前", "Main", "Snd_Now", Lanuage_Now)

If SystemTextShow = True Then '从语言文件读取
    If Lanuage_Now = 0 Or Sound_Style <= 0 Then '没选语言或者没选列表
        ReDim Preserve Sound(UBound(Sound_Name) + 3, 46)
        ReDim Preserve F_Sound(0 To 2)
        F_Sound(0) = ".Default"
        F_Sound(1) = "Explorer"
        F_Sound(2) = "sapisvr"
        Rem Windows
        Sound(0, 0) = "ChangeTheme": Sound(1, 0) = "Windows更改主题"
        Sound(0, 1) = "WindowsLogoff": Sound(1, 1) = "Windows注销"
        Sound(0, 2) = "WindowsUAC": Sound(1, 2) = "Windows用户账户控制"
        Sound(0, 3) = "WindowsLogon": Sound(1, 3) = "Windows登陆"
        Sound(0, 4) = "SystemHand": Sound(1, 4) = "关键性停止"
        Sound(0, 5) = "Close": Sound(1, 5) = "关闭程序"
        Sound(0, 6) = "RestoreUp": Sound(1, 6) = "向上还原"
        Sound(0, 7) = "RestoreDown": Sound(1, 7) = "向下还原"
        Sound(0, 8) = "MenuPopup": Sound(1, 8) = "弹出菜单"
        Sound(0, 9) = "SystemExclamation": Sound(1, 9) = "感叹号"
        Sound(0, 10) = "PrintComplete": Sound(1, 10) = "打印结束"
        Sound(0, 11) = "Open": Sound(1, 11) = "打开程序"
        Sound(0, 12) = "FaxBeep": Sound(1, 12) = "新传真通知"
        Sound(0, 13) = "MailBeep": Sound(1, 13) = "新邮件通知"
        Sound(0, 14) = "SystemAsterisk": Sound(1, 14) = "星号"
        Sound(0, 15) = "ShowBand": Sound(1, 15) = "显示工具栏区"
        Sound(0, 16) = "Maximize": Sound(1, 16) = "最大化"
        Sound(0, 17) = "Minimize": Sound(1, 17) = "最小化"
        Sound(0, 18) = "LowBatteryAlarm": Sound(1, 18) = "电池不足警报"
        Sound(0, 19) = "CriticalBatteryAlarm": Sound(1, 19) = "电池严重短缺警报"
        Sound(0, 20) = "AppGPFault": Sound(1, 20) = "程序出错"
        Sound(0, 21) = "SystemNotification": Sound(1, 21) = "系统通知"
        Sound(0, 22) = "MenuCommand": Sound(1, 22) = "菜单命令"
        Sound(0, 23) = "DeviceDisconnect": Sound(1, 23) = "设备中断连接"
        Sound(0, 24) = "DeviceFail": Sound(1, 24) = "设备未能连接"
        Sound(0, 25) = "DeviceConnect": Sound(1, 25) = "设备连接"
        Sound(0, 26) = "SystemExit": Sound(1, 26) = "退出Windows"
        Sound(0, 27) = "CCSelect": Sound(1, 27) = "选择"
        Sound(0, 28) = "SystemQuestion": Sound(1, 28) = "问题"
        Sound(0, 29) = ".Default": Sound(1, 29) = "默认响声"
        Rem Windows资源管理器
        Sound(0, 30) = "FaxError": Sound(1, 30) = "传真错误"
        Sound(0, 31) = "SecurityBand": Sound(1, 31) = "信息栏"
        Sound(0, 32) = "Navigating": Sound(1, 32) = "启动导航"
        Sound(0, 33) = "ActivatingDocument": Sound(1, 33) = "完成导航"
        Sound(0, 34) = "SearchProviderDiscovered": Sound(1, 34) = "已发现搜索提供程序"
        Sound(0, 35) = "FeedDiscovered": Sound(1, 35) = "已发现源"
        Sound(0, 36) = "FaxSent": Sound(1, 36) = "已发送传真"
        Sound(0, 37) = "FaxLineRings": Sound(1, 37) = "拨入电话"
        Sound(0, 38) = "EmptyRecycleBin": Sound(1, 38) = "清空回收站"
        Sound(0, 39) = "MoveMenuItem": Sound(1, 39) = "移动菜单项"
        Sound(0, 40) = "BlockedPopup": Sound(1, 40) = "阻止的弹出窗口"
        Rem Windows语音识别
        Sound(0, 41) = "HubOffSound": Sound(1, 41) = "关闭"
        Sound(0, 42) = "HubOnSound": Sound(1, 42) = "启用"
        Sound(0, 43) = "DisNumbersSound": Sound(1, 43) = "消除歧义号码"
        Sound(0, 44) = "PanelSound": Sound(1, 44) = "消除歧义面板"
        Sound(0, 45) = "HubSleepSound": Sound(1, 45) = "睡眠"
        Sound(0, 46) = "MisrecoSound": Sound(1, 46) = "误识别"
        '如果采用从系统读取
        'For i = 0 To 46
        '    Sound(1, i) = Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\EventLabels\" & Sound(0, i), "DispFileName"))
        'Next i
        
        Set Root = .Add(, , "F_.Default", "Windows", 1)
        For i = 0 To 29
            .Add "F_.Default", tvwChild, "s" & i, Sound(1, i)
            For k = 0 To UBound(Sound_Name) '添加每种系统预设
                Sound(k + 3, i) = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\.Default\" & Sound(0, i) & "\" & Sound_Name(k), vbNullString)
            Next k
        Next i
        Set Root = .Add(, , "F_Explorer", "Windows资源管理器", 1)
        For i = 30 To 40
            .Add "F_Explorer", tvwChild, "s" & i, Sound(1, i)
            For k = 0 To UBound(Sound_Name) '添加每种系统预设
                Sound(k + 3, i) = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\Explorer\" & Sound(0, i) & "\" & Sound_Name(k), vbNullString)
            Next k
        Next i
        Set Root = .Add(, , "F_sapisvr", "Windows语言识别", 1)
        For i = 41 To 46
            .Add "F_sapisvr", tvwChild, "s" & i, Sound(1, i)
            For k = 0 To UBound(Sound_Name) '添加每种系统预设
                Sound(k + 3, i) = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\sapisvr\" & Sound(0, i) & "\" & Sound_Name(k), vbNullString)
            Next k
        Next i
    Else
    
    '从语言文件读取
    Dim m%, n%
        If Sound_Style > 0 Then
            Dim Host_Snd_Name As String
            Host_Snd_Name = GetFromIni("Sounds", "List" & Sound_Style, Lanuages(Lanuage_Now)) '获得当前选择的音效列表名
            For i = 1 To 32767 '循环那个列表的父节点个数
                ReDim Preserve F_Sound(0 To i - 1)
                If GetFromIni(Host_Snd_Name & "_Sounds" & i, "Name", Lanuages(Lanuage_Now)) <> "" Then
                    F_Sound(i - 1) = GetFromIni(Host_Snd_Name & "_Sounds" & i, "Name", Lanuages(Lanuage_Now))
                    If GetFromIni(Host_Snd_Name & "_Sounds" & i, "DisplayName", Lanuages(Lanuage_Now)) <> "" Then
                        Set Root = .Add(, , "F_" & F_Sound(i - 1), GetFromIni(Host_Snd_Name & "_Sounds" & i, "DisplayName", Lanuages(Lanuage_Now)), 1)
                    Else
                        Set Root = .Add(, , "F_" & F_Sound(i - 1), F_Sound(i - 1), 1)
                    End If
                    
                    For j = 1 To 32767 '循环子节点个数
                        
                        If GetFromIni(Host_Snd_Name & "_Sounds" & i, "Snd" & j, Lanuages(Lanuage_Now)) <> "" Then
                        ReDim Preserve Sound(UBound(Sound_Name) + 3, 0 To n)
                        ReDim Preserve S_Sound(0 To j)
                            S_Sound(j - 1) = GetFromIni(Host_Snd_Name & "_Sounds" & i, "Snd" & j, Lanuages(Lanuage_Now))
                            Sound(0, n) = Left(S_Sound(j - 1), InStr(S_Sound(j - 1), "\") - 1)
                            Sound(1, n) = Mid(S_Sound(j - 1), InStr(S_Sound(j - 1), "\") + 1)
                            .Add "F_" & F_Sound(i - 1), tvwChild, "s" & n, Sound(1, n)
                            Root.Sorted = True '对子节点进行排序
                            For k = 0 To UBound(Sound_Name) '添加每种系统预设
                                Sound(k + 3, n) = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i - 1) & "\" & Sound(0, n) & "\" & Sound_Name(k), vbNullString)
                            Next k
                        n = n + 1
                        Else
                            Exit For
                        End If
                    Next j
                Else
                    ReDim Preserve F_Sound(0 To i - 2) '如果空了的话就取消掉增加的这个空的
                    Exit For
                End If
            Next i
        End If

    End If
Else '从系统读取
    n = 0
    Call GetAllKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps", F_Sound)
    For i = 0 To UBound(F_Sound) '父节点循环
        If GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i), "DispFileName") <> "" Then
            Set Root = .Add(, , "F_" & F_Sound(i), Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i), "DispFileName")), 1)
        ElseIf GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i), vbNullString) <> "" Then
            Set Root = .Add(, , "F_" & F_Sound(i), Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i), vbNullString)), 1)
        Else
            Set Root = .Add(, , "F_" & F_Sound(i), F_Sound(i))
        End If
    
        Call GetAllKey(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i), S_Sound)
        For j = 0 To UBound(S_Sound) '子节点循环
            ReDim Preserve Sound(UBound(Sound_Name) + 3, 0 To n)
            Sound(0, n) = S_Sound(j)
            If Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\EventLabels\" & Sound(0, n), "DispFileName")) <> "" Then
                Sound(1, n) = Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\EventLabels\" & Sound(0, n), "DispFileName"))
            ElseIf Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\EventLabels\" & Sound(0, n), vbNullString)) <> "" Then
                Sound(1, n) = Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\EventLabels\" & Sound(0, n), vbNullString))
            Else
                Sound(1, n) = Sound(0, n)
            End If
            .Add "F_" & F_Sound(i), tvwChild, "s" & n, Sound(1, n) '加入子节点
            Root.Sorted = True '对子节点进行排序
            For k = 0 To UBound(Sound_Name) '添加每种系统预设
                Sound(k + 3, n) = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & S_Sound(j) & "\" & Sound_Name(k), vbNullString)
            Next k
            n = n + 1
        Next j
    Next i
End If

End With
'给系统颜色添加数组，英文名
SysColors(0, 0) = "Scrollbar"
SysColors(1, 0) = "Background"
SysColors(2, 0) = "ActiveTitle"
SysColors(3, 0) = "InactiveTitle"
SysColors(4, 0) = "Menu"
SysColors(5, 0) = "Window"
SysColors(6, 0) = "WindowFrame"
SysColors(7, 0) = "MenuText"
SysColors(8, 0) = "WindowText"
SysColors(9, 0) = "TitleText"
SysColors(10, 0) = "ActiveBorder"
SysColors(11, 0) = "InactiveBorder"
SysColors(12, 0) = "AppWorkspace"
SysColors(13, 0) = "Hilight"
SysColors(14, 0) = "HilightText"
SysColors(15, 0) = "ButtonFace"
SysColors(16, 0) = "ButtonShadow"
SysColors(17, 0) = "GrayText"
SysColors(18, 0) = "ButtonText"
SysColors(19, 0) = "InactiveTitleText"
SysColors(20, 0) = "ButtonHilight"
SysColors(21, 0) = "ButtonDkShadow"
SysColors(22, 0) = "ButtonLight"
SysColors(23, 0) = "InfoText"
SysColors(24, 0) = "InfoWindow"
SysColors(25, 0) = "ButtonAlternateFace"
SysColors(26, 0) = "HotTrackingColor"
SysColors(27, 0) = "GradientActiveTitle"
SysColors(28, 0) = "GradientInactiveTitle"
SysColors(29, 0) = "MenuHilight"
SysColors(30, 0) = "MenuBar"
'给系统颜色添加数组，Windows 经典
SysColors(0, 2) = "212 208 200"
SysColors(1, 2) = "58 110 165"
SysColors(2, 2) = "10 36 106"
SysColors(3, 2) = "128 128 128"
SysColors(4, 2) = "212 208 200"
SysColors(5, 2) = "255 255 255"
SysColors(6, 2) = "0 0 0"
SysColors(7, 2) = "0 0 0"
SysColors(8, 2) = "0 0 0"
SysColors(9, 2) = "255 255 255"
SysColors(10, 2) = "212 208 200"
SysColors(11, 2) = "212 208 200"
SysColors(12, 2) = "128 128 128"
SysColors(13, 2) = "10 36 106"
SysColors(14, 2) = "255 255 255"
SysColors(15, 2) = "212 208 200"
SysColors(16, 2) = "128 128 128"
SysColors(17, 2) = "128 128 128"
SysColors(18, 2) = "0 0 0"
SysColors(19, 2) = "212 208 200"
SysColors(20, 2) = "255 255 255"
SysColors(21, 2) = "64 64 64"
SysColors(22, 2) = "212 208 200"
SysColors(23, 2) = "0 0 0"
SysColors(24, 2) = "255 255 225"
SysColors(25, 2) = "181 181 181"
SysColors(26, 2) = "0 0 128"
SysColors(27, 2) = "166 202 240"
SysColors(28, 2) = "192 192 192"
SysColors(29, 2) = "10 36 106"
SysColors(30, 2) = "212 208 200"
'给系统颜色添加数组高对比度 #1
SysColors(0, 3) = "0 0 0"
SysColors(1, 3) = "0 0 0"
SysColors(2, 3) = "0 0 255"
SysColors(3, 3) = "0 255 255"
SysColors(4, 3) = "0 0 0"
SysColors(5, 3) = "0 0 0"
SysColors(6, 3) = "255 255 255"
SysColors(7, 3) = "255 255 255"
SysColors(8, 3) = "255 255 0"
SysColors(9, 3) = "255 255 255"
SysColors(10, 3) = "0 0 255"
SysColors(11, 3) = "0 255 255"
SysColors(12, 3) = "0 0 0"
SysColors(13, 3) = "0 128 0"
SysColors(14, 3) = "255 255 255"
SysColors(15, 3) = "0 0 0"
SysColors(16, 3) = "128 128 128"
SysColors(17, 3) = "0 255 0"
SysColors(18, 3) = "255 255 255"
SysColors(19, 3) = "0 0 0"
SysColors(20, 3) = "192 192 192"
SysColors(21, 3) = "255 255 255"
SysColors(22, 3) = "255 255 255"
SysColors(23, 3) = "255 255 0"
SysColors(24, 3) = "0 0 0"
SysColors(25, 3) = "192 192 192"
SysColors(26, 3) = "128 128 255"
SysColors(27, 3) = "0 0 255"
SysColors(28, 3) = "0 255 255"
SysColors(29, 3) = "0 128 0"
SysColors(30, 3) = "0 0 0"
'给系统颜色添加数组高对比度 #2
SysColors(0, 4) = "0 0 0"
SysColors(1, 4) = "0 0 0"
SysColors(2, 4) = "0 255 255"
SysColors(3, 4) = "0 0 255"
SysColors(4, 4) = "0 0 0"
SysColors(5, 4) = "0 0 0"
SysColors(6, 4) = "255 255 255"
SysColors(7, 4) = "0 255 0"
SysColors(8, 4) = "0 255 0"
SysColors(9, 4) = "0 0 0"
SysColors(10, 4) = "0 255 255"
SysColors(11, 4) = "0 0 255"
SysColors(12, 4) = "255 255 255"
SysColors(13, 4) = "0 0 255"
SysColors(14, 4) = "255 255 255"
SysColors(15, 4) = "0 0 0"
SysColors(16, 4) = "128 128 128"
SysColors(17, 4) = "192 192 192"
SysColors(18, 4) = "0 255 0"
SysColors(19, 4) = "255 255 255"
SysColors(20, 4) = "192 192 192"
SysColors(21, 4) = "255 255 255"
SysColors(22, 4) = "255 255 255"
SysColors(23, 4) = "0 0 0"
SysColors(24, 4) = "255 255 0"
SysColors(25, 4) = "192 192 192"
SysColors(26, 4) = "128 128 255"
SysColors(27, 4) = "0 255 255"
SysColors(28, 4) = "0 0 255"
SysColors(29, 4) = "0 0 255"
SysColors(30, 4) = "0 0 0"
'给系统颜色添加数组高对比度黑色
SysColors(0, 5) = "0 0 0"
SysColors(1, 5) = "0 0 0"
SysColors(2, 5) = "128 0 128"
SysColors(3, 5) = "0 128 0"
SysColors(4, 5) = "0 0 0"
SysColors(5, 5) = "0 0 0"
SysColors(6, 5) = "255 255 255"
SysColors(7, 5) = "255 255 255"
SysColors(8, 5) = "255 255 255"
SysColors(9, 5) = "255 255 255"
SysColors(10, 5) = "255 255 0"
SysColors(11, 5) = "0 128 0"
SysColors(12, 5) = "0 0 0"
SysColors(13, 5) = "128 0 128"
SysColors(14, 5) = "255 255 255"
SysColors(15, 5) = "0 0 0"
SysColors(16, 5) = "128 128 128"
SysColors(17, 5) = "0 255 0"
SysColors(18, 5) = "255 255 255"
SysColors(19, 5) = "255 255 255"
SysColors(20, 5) = "192 192 192"
SysColors(21, 5) = "255 255 255"
SysColors(22, 5) = "255 255 255"
SysColors(23, 5) = "255 255 255"
SysColors(24, 5) = "0 0 0"
SysColors(25, 5) = "192 192 192"
SysColors(26, 5) = "128 128 255"
SysColors(27, 5) = "128 0 128"
SysColors(28, 5) = "0 128 0"
SysColors(29, 5) = "128 0 128"
SysColors(30, 5) = "0 0 0"
'给系统颜色添加数组高对比度白色
SysColors(0, 6) = "255 255 255"
SysColors(1, 6) = "255 255 255"
SysColors(2, 6) = "0 0 0"
SysColors(3, 6) = "255 255 255"
SysColors(4, 6) = "255 255 255"
SysColors(5, 6) = "255 255 255"
SysColors(6, 6) = "0 0 0"
SysColors(7, 6) = "0 0 0"
SysColors(8, 6) = "0 0 0"
SysColors(9, 6) = "255 255 255"
SysColors(10, 6) = "128 128 128"
SysColors(11, 6) = "192 192 192"
SysColors(12, 6) = "128 128 128"
SysColors(13, 6) = "0 0 0"
SysColors(14, 6) = "255 255 255"
SysColors(15, 6) = "255 255 255"
SysColors(16, 6) = "128 128 128"
SysColors(17, 6) = "0 128 0"
SysColors(18, 6) = "0 0 0"
SysColors(19, 6) = "0 0 0"
SysColors(20, 6) = "192 192 192"
SysColors(21, 6) = "0 0 0"
SysColors(22, 6) = "192 192 192"
SysColors(23, 6) = "0 0 0"
SysColors(24, 6) = "255 255 255"
SysColors(25, 6) = "192 192 192"
SysColors(26, 6) = "0 0 0"
SysColors(27, 6) = "0 0 0"
SysColors(28, 6) = "255 255 255"
SysColors(29, 6) = "0 0 0"
SysColors(30, 6) = "255 255 255"

With Main
    If .TreeView_Sound.Nodes.count > 0 Then
        '使全部展开
        For i = 1 To .TreeView_Sound.Nodes.count
            .TreeView_Sound.Nodes(i).Expanded = True '展开所有节点
        Next i
        '默认选中第一个
        .TreeView_Sound.Nodes(1).Selected = True
    End If
    '图标的系统名字
    For i = 0 To 5
        .Label_icon(i).Caption = Get_dll_text(SysIco(i, 2))
    Next
    '对鼠标们进行排列
    For i = 0 To 14
        .Cur_BG(i).Top = 510 * i + 20
        .Cur_BG(i).Left = 20
        .Cur_BG(i).BackColor = &HFFFFFF
        .Pic_Cur(i).Top = 510 * i + 20
        .Pic_Cur(i).Left = 4220
        .Pic_Cur(i).BorderStyle = 0
        .Cur_BG(i).Caption = vbCrLf & "    " & Get_dll_text(SysCur(i, 3)) '鼠标们的名字
    Next
    .Frame_Mouse.Width = 320 * 15
    .Frame_Mouse.Height = 1000 * 15
    .Frame_Mouse.BackColor = &HFFFFFF
    
    Dim ComboRoot As ComboItem
    With .ImageCombo_paper_style.ComboItems
        Dim Paper_Style_Name(1 To 5) As String
        If SystemTextShow = True Or System_Ver < 6 Then
            If Lanuage_Now = 0 Then
                Paper_Style_Name(1) = "填充"
                Paper_Style_Name(2) = "适应"
                Paper_Style_Name(3) = "拉伸"
                Paper_Style_Name(4) = "平铺"
                Paper_Style_Name(5) = "居中"
            Else
                For i = 0 To 4
                    Paper_Style_Name(i + 1) = Load_Lanuage("Paper Change Time" & i, "Main", "Paper_Style_Name" & i, Lanuage_Now)
                Next
            End If
        Else
            For i = 0 To 4
                Paper_Style_Name(i + 1) = Get_dll_text("@themecpl.dll,-" & 504 + i)
            Next
        End If
    .Clear '先清除
    Set ComboRoot = .Add(1, "填充", Paper_Style_Name(1), 1) '10
    Set ComboRoot = .Add(2, "适应", Paper_Style_Name(2), 2) '6
    Set ComboRoot = .Add(3, "拉伸", Paper_Style_Name(3), 3) '2
    Set ComboRoot = .Add(4, "平铺", Paper_Style_Name(4), 4) '0，但是TileWallpaper=1
    Set ComboRoot = .Add(5, "居中", Paper_Style_Name(5), 5) '0，但是TileWallpaper=0
    '只要选TileWallpaper就是平铺
    End With
    .ImageCombo_paper_style.ComboItems(1).Selected = True
    '壁纸切换时间
    With .Combo_Paper_Change_Time
        Dim Change_Time_name(14) As String
        If SystemTextShow = True Or System_Ver < 6 Then
            If Lanuage_Now = 0 Then
                Change_Time_name(0) = "10秒"
                Change_Time_name(1) = "30秒"
                Change_Time_name(2) = "1分钟"
                Change_Time_name(3) = "3分钟"
                Change_Time_name(4) = "5分钟"
                Change_Time_name(5) = "10分钟"
                Change_Time_name(6) = "15分钟"
                Change_Time_name(7) = "20分钟"
                Change_Time_name(8) = "30分钟"
                Change_Time_name(9) = "1小时"
                Change_Time_name(10) = "2小时"
                Change_Time_name(11) = "3小时"
                Change_Time_name(12) = "6小时"
                Change_Time_name(13) = "12小时"
                Change_Time_name(14) = "一天"
            Else
                For i = 0 To 14
                    Change_Time_name(i) = Load_Lanuage("Paper Change Time" & i, "Main", "Paper_Change_Time" & i, Lanuage_Now)
                Next
            End If
        Else
            For i = 0 To 14
                Change_Time_name(i) = Get_dll_text("@themecpl.dll,-" & 509 + i)
            Next
        End If
    .Clear '先清除
        For i = 0 To 14
            .AddItem Change_Time_name(i), i
        Next
        .ListIndex = 8
    End With
    
    
    '经典风格预设
    With .ImageCombo_Classic_Style.ComboItems
        Dim Classic_Style_name(1 To 6) As String
        If SystemTextShow = True Then
            If Lanuage_Now = 0 Then
                Classic_Style_name(1) = "自定义"
                Classic_Style_name(2) = "Windows 经典"
                Classic_Style_name(3) = "高对比度 #1"
                Classic_Style_name(4) = "高对比度 #2"
                Classic_Style_name(5) = "高对比黑色"
                Classic_Style_name(6) = "高对比白色"
            Else
                For i = 0 To 5
                    Classic_Style_name(i + 1) = Load_Lanuage("Classic Style name" & i, "Main", "Classic_Style_name" & i, Lanuage_Now)
                Next
            End If
        Else
        Classic_Style_name(1) = Load_Lanuage("Custom", "Main", "Classic_Style_name0", Lanuage_Now)
        Classic_Style_name(2) = Get_dll_text("@themeui.dll,-854")
        Classic_Style_name(3) = Get_dll_text("@themeui.dll,-850")
        Classic_Style_name(4) = Get_dll_text("@themeui.dll,-851")
        Classic_Style_name(5) = Get_dll_text("@themeui.dll,-852")
        Classic_Style_name(6) = Get_dll_text("@themeui.dll,-853")
        End If
    .Clear '先清除
    For i = 1 To 6
        Set ComboRoot = .Add(i, "C_S" & i, Classic_Style_name(i), i) '这里的key关系到XP能不能正常显示本程序
    Next
    End With
    .ImageCombo_Classic_Style.ComboItems(2).Selected = True

    '对系统颜色们进行排列
    For i = 0 To 30
        .Lable_System_Color(i).Caption = SysColors(i, 0) '设置左边的名称
        .Lable_System_Color(i).Top = 290 * i + 60
        .Lable_System_Color(i).Left = 50
        .Value_System_Color(i).Top = 290 * i + 45
        .Value_System_Color(i).Left = 1970
    Next
    .Frame_System_Color.Width = 209 * 15
    .Frame_System_Color.Height = 900 * 15
End With

Call Refresh_Theme
End Sub

Public Sub Refresh_Theme()
Dim i As Integer
Dim Root As Node
'添加主题列表
    Main.TreeView_Theme.Nodes.Clear '清除以前的节点
    Main.ImageList_Theme.ListImages.Clear '清除以前的图片
    'Erase Theme1 '先清空数组
    'Erase Theme2 '先清空数组
    Set Theme1 = New Collection
    Set Theme2 = New Collection
    
    'Dim UBound_temp1 As Long, UBound_temp2 As Long
    'Dim Theme1_url_temp As String, Theme2_url_temp As String
    'Dim Theme_Files_url(2) As String, Theme_Reg_url(1) As String
    Dim Theme_Files_url As New Collection, Theme_Reg_url As New Collection
    
'    Theme_Files_url(0) = url_to_N("%SystemRoot%\Resources\Themes")
    Theme_Files_url.Add url_to_N("%SystemRoot%\Resources")
    Theme_Files_url.Add Environ("LocalAppData") & "\Microsoft\Windows\Themes"
    Theme_Files_url.Add url_to_N("%SystemRoot%\Globalization\MCT\")
'    Theme_Files_url(2) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-AU\Theme")
'    Theme_Files_url(3) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-CA\Theme")
'    Theme_Files_url(4) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-GB\Theme")
'    Theme_Files_url(5) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-US\Theme")
'    Theme_Files_url(6) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-ZA\Theme")
'    Theme_Files_url(7) = url_to_N("%SystemRoot%\Globalization\MCT\MCT-CN\Theme")
    Theme_Reg_url.Add "Software\Microsoft\Windows\CurrentVersion\Themes\InstalledThemes\MCT"
    Theme_Reg_url.Add "Software\Microsoft\Windows\CurrentVersion\Themes\InstalledThemes\SQM"
    
    'ReDim Preserve Theme1(0) '先给theme1一个0数组
    
    '注册表注册部分
    Dim asTemp
    'For i = 0 To UBound(Theme_Reg_url)
    For i = 1 To Theme_Reg_url.count
        If RegOpenKey(HKEY_CURRENT_USER, Theme_Reg_url(i), asTemp) = 0 Then
            Call GetAllValue(HKEY_CURRENT_USER, Theme_Reg_url(i), Theme2)
            'UBound_temp2 = UBound(Theme2)
            'For j = 0 To UBound_temp2
            For j = 1 To Theme2.count
                If NewTheme(Theme2(j)) Then
                    'UBound_temp1 = UBound(Theme1)
                    'ReDim Preserve Theme1(0 To UBound_temp1 + 1) '扩展1
                    'Theme1(UBound_temp1 + 1) = Theme2(j)
                    Theme1.Add Theme2(j)
                End If
            Next j
        End If
        'Erase Theme2 '先清空数组
        Set Theme2 = New Collection
    Next i
    
    '特殊路径部分
    'For i = 0 To UBound(Theme_Files_url)
    For i = 1 To Theme_Files_url.count
        If Dir(Theme_Files_url(i), vbDirectory + vbHidden + vbSystem) <> "" Then
            'ReDim Preserve Theme2(0)
            Call GetFileNameRecursion(Theme_Files_url(i), "theme", Theme2) '获取安装theme文件列表
        'On Error GoTo Theme1
            'UBound_temp1 = UBound(Theme1)
                'UBound_temp2 = UBound(Theme2)
                'For j = 0 To UBound_temp2
                For j = 1 To Theme2.count
                    If NewTheme(Theme2(j)) Then
                        'UBound_temp1 = UBound(Theme1)
                        'ReDim Preserve Theme1(0 To UBound_temp1 + 1) '扩展1
                        'Theme1(UBound_temp1 + 1) = Theme2(j)
                        Theme1.Add Theme2(j)
                    End If
                Next j
'            ReDim Preserve Theme1(0 To (UBound_temp1 + UBound_temp2)) '扩展1
'            For j = 1 To UBound_temp2
'                Theme1(UBound_temp1 + j) = Theme2(j)
'            Next j
        End If
        'Erase Theme2 '先清空数组
        Set Theme2 = New Collection
'Theme1:
    Next i

'For j = 0 To UBound(Theme1)
'    Debug.Print Theme1(j)
'Next j

    '将所有图片载入imagelist
    Dim Paper_Url As String, Papers_Folder As String, BackGround_Color As String, BGColor() As String
    Dim CColor As String
    Dim Graphics As Long, penH As Long, brushH As Long
    
    'Dim Papers_Url() As String
    Dim Papers_Url As Collection

'    Dim cIcon As New cAniCursor '图标
'    cIcon.LoadFromFile url_to_N("%SystemRoot%\system32\SHELL32.dll,3")
'    cIcon.Draw Main.Picture_paper_TEMP.hDC, 0, 0, 68, 68, Main.Picture_paper_TEMP.BackColor
'    Main.ImageList_Theme.ListImages.Add , , Main.Picture_paper_TEMP.Image
    
    Dim rootPathPapers As Collection
    'For i = 1 To UBound(Theme1)
    For i = 1 To Theme1.count
        Main.Picture_paper_TEMP.Cls
        Paper_Url = url_to_N(GetFromIni("Control Panel\Desktop", "Wallpaper", Theme1(i)))
        Papers_Folder = url_to_N(GetFromIni("Slideshow", "ImagesRootPath", Theme1(i)))
        
        Set rootPathPapers = New Collection
        If Len(Papers_Folder) > 0 And Dir(Main.url_paper_files.text, vbDirectory + vbHidden + vbSystem) <> "" Then
            Call GetFileName(Main.url_paper_files.text, "bmp,jpg,jpeg,gif,png", rootPathPapers) '获取该文件夹下图片文件列表
        End If
        If rootPathPapers.count = 0 Or Papers_Folder = "" And url_to_N(GetFromIni("Slideshow", "ImagesRootPIDL", Theme1(i))) <> "" Then
            Papers_Folder = Left(Paper_Url, InStrRev(Paper_Url, "\"))
        End If
        If Papers_Folder <> "" And Len(Dir(Papers_Folder, vbDirectory)) > 0 Then
            'Erase Papers_Url '先清空数组
            Set Papers_Url = New Collection '先清空数组
            Call GetFileName(Papers_Folder, "bmp,jpg,jpeg,gif,png", Papers_Url) '获取安装theme文件列表
        'On Error GoTo Paper1
            'If UBound(Papers_Url) > 0 Then
            If Papers_Url.count > 0 Then
                'If UBound(Papers_Url) > 2 Then
                If Papers_Url.count > 2 Then
                    For j = 1 To 3
                        Call PaintPng2(Papers_Url(j), Main.Picture_paper_TEMP.hDC, pWidth - 20, pHeight - 20, 20 - 10 * (j - 1), 10 * (j - 1))
                    Next
                Else
                    'For j = 1 To UBound(Papers_Url)
                    For j = 1 To Papers_Url.count
                        'Call PaintPng2(Papers_Url(j), Main.Picture_paper_TEMP.hDC, pWidth - (UBound(Papers_Url) - 1) * 10, pHeight - (UBound(Papers_Url) - 1) * 10, (UBound(Papers_Url) - 1) * 10 - 10 * (j - 1), 10 * (j - 1))
                        Call PaintPng2(Papers_Url(j), Main.Picture_paper_TEMP.hDC, pWidth - (Papers_Url.count - 1) * 10, pHeight - (Papers_Url.count - 1) * 10, (Papers_Url.count - 1) * 10 - 10 * (j - 1), 10 * (j - 1))
                    Next
                End If
            End If
'Paper1:
        ElseIf Paper_Url <> "" And Len(Dir(Paper_Url)) > 0 Then
            Call PaintPng2(Paper_Url, Main.Picture_paper_TEMP.hDC, pWidth, pHeight)
        Else
            BackGround_Color = GetFromIni("Control Panel\Colors", "Background", Theme1(i))
            If Len(BackGround_Color) > 0 Then
                BGColor = Split(BackGround_Color, " ")
                If UBound(BGColor) >= 2 Then
                    BackGround_Color = RGB(CByte(BGColor(0)), CByte(BGColor(1)), CByte(BGColor(2)))
                    Main.Picture_paper_TEMP.Line (0, 0)-Step(Main.Picture_paper_TEMP.ScaleWidth, Main.Picture_paper_TEMP.ScaleHeight), BackGround_Color, BF
                End If
            End If
        End If
        
        CColor = GetFromIni("VisualStyles", "ColorizationColor", Theme1(i))
        If CColor <> "" Then
            CColor = text_to_color(CColor)

            InitGDIPlus '初始化GDI+

            GdipCreateFromHDC Main.Picture_paper_TEMP.hDC, Graphics '创建图像

            GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias '去锯齿
            
            'GdipFillRectangleI Graphics, brushH, pWidth - pHeight * 0.5, pHeight * 0.5, pHeight * 0.5, pHeight * 0.5 '填充
            'GdipDrawRectangleI Graphics, penH, pWidth - pHeight * 0.5, pHeight * 0.5, pHeight * 0.5, pHeight * 0.5  '画框
            
            Dim tPoints(24) As POINTF
            Dim bLength As Integer, cRadius As Integer
            Dim osx As Integer, osy As Integer, sc As Single
            bLength = 40 '宽度
            cRadius = 8 '圆角半径
            sc = 1 '放大比率
            osx = pWidth - bLength * sc - 2 'X偏移
            osy = pHeight - bLength * sc - 2 'Y偏移
            
            '储存曲线点
            tPoints(0).x = osx + cRadius * sc: tPoints(0).y = osy + 0 * sc
            tPoints(1).x = osx + bLength / 2 * sc: tPoints(1).y = osy + 0 * sc
            tPoints(2).x = osx + (bLength - cRadius) * sc: tPoints(2).y = osy + 0 * sc
            tPoints(3).x = osx + (bLength - cRadius) * sc: tPoints(3).y = osy + 0 * sc
            tPoints(4).x = osx + bLength * sc: tPoints(4).y = osy + 0 * sc
            tPoints(5).x = osx + bLength * sc: tPoints(5).y = osy + cRadius * sc
            tPoints(6).x = osx + bLength * sc: tPoints(6).y = osy + cRadius * sc
            tPoints(7).x = osx + bLength * sc: tPoints(7).y = osy + bLength / 2 * sc
            tPoints(8).x = osx + bLength * sc: tPoints(8).y = osy + (bLength - cRadius) * sc
            tPoints(9).x = osx + bLength * sc: tPoints(9).y = osy + (bLength - cRadius) * sc
            tPoints(10).x = osx + bLength * sc: tPoints(10).y = osy + bLength * sc
            tPoints(11).x = osx + (bLength - cRadius) * sc: tPoints(11).y = osy + bLength * sc
            tPoints(12).x = osx + (bLength - cRadius) * sc: tPoints(12).y = osy + bLength * sc
            tPoints(13).x = osx + bLength / 2 * sc: tPoints(13).y = osy + bLength * sc
            tPoints(14).x = osx + cRadius * sc: tPoints(14).y = osy + bLength * sc
            tPoints(15).x = osx + cRadius * sc: tPoints(15).y = osy + bLength * sc
            tPoints(16).x = osx + 0 * sc: tPoints(16).y = osy + bLength * sc
            tPoints(17).x = osx + 0 * sc: tPoints(17).y = osy + (bLength - cRadius) * sc
            tPoints(18).x = osx + 0 * sc: tPoints(18).y = osy + (bLength - cRadius) * sc
            tPoints(19).x = osx + 0 * sc: tPoints(19).y = osy + bLength / 2 * sc
            tPoints(20).x = osx + 0 * sc: tPoints(20).y = osy + cRadius * sc
            tPoints(21).x = osx + 0 * sc: tPoints(21).y = osy + cRadius * sc
            tPoints(22).x = osx + 0 * sc: tPoints(22).y = osy + 0 * sc
            tPoints(23).x = osx + cRadius * sc: tPoints(23).y = osy + 0 * sc
            tPoints(24).x = tPoints(0).x: tPoints(24).y = tPoints(0).y
            
            
            GdipCreatePen1 &H99000000, 1, UnitPixel, penH '创建笔
'            GdipCreateSolidFill x16_to_x10(Mid$(CColor, 3, 8)), brushH '创建刷子
                
            Dim p1 As POINTL, p2 As POINTL
            p1.x = osx + bLength * sc / 2.2
            p1.y = osy + bLength * sc / 2.3
            p2.x = osx + bLength * (sc)
            p2.y = osy + bLength * (sc)
            GdipCreateLineBrushI p1, p2, x16_to_x10(Mid$(CColor, 3, 8)), &H66FFFFFF, WrapModeTileFlipY, brushH '创建渐变刷子
            
            Dim rPath  As Long
            GdipCreatePath FillModeWinding, rPath '新建路径
            GdipAddPathBeziers rPath, tPoints(0), 25 '将之前设定的路径点添加进来
            GdipFillPath Graphics, brushH, rPath '填充路径
            GdipDrawPath Graphics, penH, rPath '画线
            GdipDeleteGraphics Graphics '删除图像
            GdipDeletePen penH '删除笔
            GdipDeleteBrush brushH '删除刷子
            GdipDeletePath rPath '删除路径
            Main.Picture_paper_TEMP.Refresh
            TerminateGDIPlus

        End If
        Main.ImageList_Theme.ListImages.Add , , Main.Picture_paper_TEMP.Image
    Next i
    '添加listview节点
    Dim Theme_Neme As String
    Dim ThemeFolder As String, fnum As Integer, hadF As Boolean
    With Main.TreeView_Theme.Nodes
        'For i = 1 To UBound(Theme1)
        For i = 1 To Theme1.count
            If GetFromIni("Theme", "DisplayName", Theme1(i)) <> "" Then
                Theme_Neme = Get_dll_text(GetFromIni("Theme", "DisplayName", Theme1(i)))
            Else
                Theme_Neme = Mid(Theme1(i), InStrRev(Theme1(i), "\") + 1, InStrRev(Theme1(i), ".") - InStrRev(Theme1(i), "\") - 1)
            End If
            ThemeFolder = Left(Theme1(i), InStrRev(Theme1(i), "\"))
            hadF = False
            For j = 1 To .count
                If .item(j).text = ThemeFolder Then
                    hadF = True
                    Exit For
                End If
            Next
'            If hadF = False Then
'                fnum = fnum + 1
'                Set Root = .Add(, , "f" & fnum, ThemeFolder, 1)
'            End If
'            Call .Add("f" & fnum, tvwChild, "t" & i, Theme_Neme, i + 1)
            Call .Add(, , "t" & i, Theme_Neme, i)
        Next
        If .count > 0 Then
        '使全部展开
            For i = 1 To .count
                .item(i).Expanded = True '展开所有节点
            Next i
        End If
    End With
'    Main.TreeView_Theme.Nodes(1).Selected = True
'Theme3:
    '下标错误（没有图片）
    

End Sub
Private Function NewTheme(ThemeURL)
    NewTheme = True
    'For i = 0 To UBound(Theme1)
    For i = 1 To Theme1.count
        If Theme1(i) = ThemeURL Then NewTheme = False
    Next i
End Function
