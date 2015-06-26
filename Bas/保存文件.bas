Attribute VB_Name = "保存文件"
Option Explicit
'宽度补齐
Public Function JuZhongBuQi(ByVal old_text As String, ByVal in_text As String, ByVal between_text As String) As String
Dim i As Integer
Dim text_num_temp As Single
Dim text_temp1, text_temp2 As String
text_num_temp = (LenB(StrConv(old_text, vbFromUnicode)) - LenB(StrConv(in_text, vbFromUnicode))) / 2

If text_num_temp <= 0 Then
    text_temp1 = ""
    text_temp2 = ""
ElseIf text_num_temp <> Int(text_num_temp) Then
    For i = 1 To Fix(text_num_temp)
        text_temp1 = text_temp1 & between_text
        text_temp2 = text_temp2 & between_text
    Next
    text_temp2 = text_temp2 & between_text
ElseIf text_num_temp = Int(text_num_temp) Then
    For i = 1 To Fix(text_num_temp)
        text_temp1 = text_temp1 & between_text
        text_temp2 = text_temp2 & between_text
    Next
End If

    JuZhongBuQi = text_temp1 & in_text & text_temp2
End Function
'保存BAT
Public Sub Save_Bat(ByVal save_url As String, Optional ByVal BAT_Color_Fore As Byte = 8, Optional ByVal BAT_Color_Back As Byte = 0)
Open save_url For Output As #1

Dim i%, j%, n%
Dim BAT_TEXT As String
Dim text_num_temp As Single
Dim text_temp1, text_temp2 As String

BAT_TEXT = "@echo off" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM " & Load_Lanuage("此主题文件由枫谷主题", "Save_Theme", "Theme_From1", Lanuage_Now) & " V" & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision & Load_Lanuage("生成", "Save_Theme", "Theme_From2", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 基于Win7 HomeBasic应用主题BAT模板1.2 by枫谷剑仙" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM “★”开头的表示是一类设置，“◢↓◣”和“◥↑◤”之间的是变量" & vbCrLf & vbCrLf
BAT_TEXT = BAT_TEXT & "title " & Main.T_name_C.text & " " & Load_Lanuage("Win7主题 HomeBasic安装程序", "Save_Bat", "Title_Start", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "color " & x10_to_x16(BAT_Color_Back, 1) & x10_to_x16(BAT_Color_Fore, 1) & vbCrLf
BAT_TEXT = BAT_TEXT & "mode con cols=42 lines=24" & vbCrLf & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("请右键以管理员权限打开本程序", "Save_Bat", "Use_UAC", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■■■■■■■■■■■■■■■■■■■■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■                                    ■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■" & JuZhongBuQi("                                    ", Main.T_name_C.text, " ") & "■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■" & JuZhongBuQi("                                    ", "By" & Main.Maker_Name.text, " ") & "■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■" & JuZhongBuQi("                                    ", Main.Maker_Web_Url.text, " ") & "■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■                                    ■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ■■■■■■■■■■■■■■■■■■■■" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ========================================" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  " & Load_Lanuage("注：本程序用于Win7 HomeBasic用户应用已经安装完成了的", "Save_Bat", "Warn_SetUp_First", Lanuage_Now) & Main.T_name_C.text & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  ========================================" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo  " & Load_Lanuage("请输入对应数字，然后按回车", "Save_Bat", "Choose", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & JuZhongBuQi("                                    ", Load_Lanuage("1.  应用到系统", "Save_Bat", "Choose_1", Lanuage_Now), " ") & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & JuZhongBuQi("                                    ", Load_Lanuage("2.  退出", "Save_Bat", "Choose_2", Lanuage_Now), " ") & vbCrLf
BAT_TEXT = BAT_TEXT & "echo.   " & vbCrLf
BAT_TEXT = BAT_TEXT & "echo ※※※※※※※※※※※※※※※※※※※※" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & ":cho" & vbCrLf
BAT_TEXT = BAT_TEXT & "set choice=" & vbCrLf
BAT_TEXT = BAT_TEXT & "set /p choice= " & Load_Lanuage("请选择:", "Save_Bat", "Choose_Please", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "IF NOT" & Chr(34) & "%choice%" & Chr(34) & "==" & Chr(34) & Chr(34) & " SET choice=%choice:~0,1%" & vbCrLf
BAT_TEXT = BAT_TEXT & "if /i " & Chr(34) & "%choice%" & Chr(34) & "==" & Chr(34) & "1" & Chr(34) & " goto setup" & vbCrLf
BAT_TEXT = BAT_TEXT & "if /i " & Chr(34) & "%choice%" & Chr(34) & "==" & Chr(34) & "2" & Chr(34) & " goto exit" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("选择无效，请重新输入", "Save_Bat", "Choose_Error", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "goto cho" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf

BAT_TEXT = BAT_TEXT & ":setup" & vbCrLf
BAT_TEXT = BAT_TEXT & "cls" & vbCrLf
If Main.Maker_Introduce.text <> "" Then
BAT_TEXT = BAT_TEXT & "title " & Load_Lanuage("版权信息", "Save_Bat", "Title_Copyright", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Replace(Main.Maker_Introduce, vbCrLf, vbCrLf & "echo ") & vbCrLf '把换行符替换为一个换行符+echo
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "pause" & vbCrLf
BAT_TEXT = BAT_TEXT & "cls" & vbCrLf
End If
BAT_TEXT = BAT_TEXT & "title " & Main.T_name_C.text & " " & Load_Lanuage("正在应用", "Save_Bat", "Title_Seting", Lanuage_Now) & vbCrLf
If Main.url_paper.text <> "" Then '为空则无壁纸
    BAT_TEXT = BAT_TEXT & "REM ★改变桌面壁纸" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set 壁纸路径=" & Main.url_paper.text & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set paper=reg add " & Chr(34) & "HKEY_CURRENT_USER\Control Panel\Desktop" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%paper%" & Chr(34) & " /v Wallpaper /d " & Chr(34) & "%壁纸路径%" & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM 壁纸填充风格为：" & Main.ImageCombo_paper_style.SelectedItem.Key & vbCrLf
    BAT_TEXT = BAT_TEXT & "%paper%" & Chr(34) & " /v TileWallpaper /d " & Chr(34) & TileWallpaper_value & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%paper%" & Chr(34) & " /v WallpaperStyle /d " & Chr(34) & WallpaperStyle_value & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
End If

If Main.url_scr.text <> "" Then '为空则无屏保
    BAT_TEXT = BAT_TEXT & "REM ★改变屏保，注意屏保等待时间单位为秒" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set 屏保路径=" & Main.url_scr.text & vbCrLf
    BAT_TEXT = BAT_TEXT & "set 屏保等待时间=" & Main.scr_wait_min * 60 & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set ScreenSave=reg add " & Chr(34) & "HKEY_CURRENT_USER\Control Panel\Desktop" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%ScreenSave%" & Chr(34) & " /v SCRNSAVE.EXE /d " & Chr(34) & "%屏保路径%" & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM 屏保等待时间" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%ScreenSave%" & Chr(34) & " /v ScreenSaveTimeOut /d " & Chr(34) & "%屏保等待时间%" & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%ScreenSave%" & Chr(34) & " /v ScreenSaveActive /d " & Chr(34) & "1" & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "%ScreenSave%" & Chr(34) & " /v ScreenSaverIsSecure /d " & Chr(34) & "0" & Chr(34) & " /f" & vbCrLf
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
End If

'图标
BAT_TEXT = BAT_TEXT & "REM ★改变桌面图标" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
For i = 0 To 5
    If SysIco(i, 0) <> "" Then
        BAT_TEXT = BAT_TEXT & "set " & SysIco(i, 1) & "=" & SysIco(i, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "set " & SysIco(i, 1) & "=" & SysIco(i, 3) & vbCrLf
    End If
Next
BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
BAT_TEXT = BAT_TEXT & "set icon=reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon" & Chr(34) & " /d " & Chr(34) & "%" & SysIco(0, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon" & Chr(34) & " /d " & Chr(34) & "%" & SysIco(1, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon" & Chr(34) & " /d " & Chr(34) & "%" & SysIco(2, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon" & Chr(34) & " /v Empty /d " & Chr(34) & "%" & SysIco(3, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon" & Chr(34) & " /v Full /d " & Chr(34) & "%" & SysIco(4, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%icon%{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon" & Chr(34) & " /d " & Chr(34) & "%" & SysIco(5, 1) & "%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf
'鼠标
BAT_TEXT = BAT_TEXT & "REM ★改变鼠标" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 设置中文（或自己喜欢的其他文）的鼠标方案名字" & vbCrLf
BAT_TEXT = BAT_TEXT & "set 鼠标方案=" & Main.name_cur.text & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 设置光标文件地址" & vbCrLf
BAT_TEXT = BAT_TEXT & "set 正常选择=" & SysCur(0, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 帮助选择=" & SysCur(1, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 后台运行=" & SysCur(2, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 忙=" & SysCur(3, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 精确定位=" & SysCur(4, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 选定文本=" & SysCur(5, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 手写=" & SysCur(6, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 不可用=" & SysCur(7, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 垂直调整=" & SysCur(8, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 水平调整=" & SysCur(9, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 对角线1=" & SysCur(10, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 对角线2=" & SysCur(11, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 移动=" & SysCur(12, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 候选=" & SysCur(13, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "set 连接选择=" & SysCur(14, 0) & vbCrLf
BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
BAT_TEXT = BAT_TEXT & "set Cursor=reg add " & Chr(34) & "HKEY_CURRENT_USER\Control Panel\Cursors" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /ve /d " & Chr(34) & "%鼠标方案%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 设置当前鼠标" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v Arrow /d " & Chr(34) & "%正常选择%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v Help /d " & Chr(34) & "%帮助选择%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v AppStarting /d " & Chr(34) & "%后台运行%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v Wait /d " & Chr(34) & "%忙%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v Crosshair /d " & Chr(34) & "%精确定位%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v IBeam /d " & Chr(34) & "%选定文本%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v NWPen /d " & Chr(34) & "%手写%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v No /d " & Chr(34) & "%不可用%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v SizeNS /d " & Chr(34) & "%垂直调整%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v SizeWE /d " & Chr(34) & "%水平调整%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v SizeNWSE /d " & Chr(34) & "%对角线1%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v SizeNESW /d " & Chr(34) & "%对角线2%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v SizeAll /d " & Chr(34) & "%移动%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v UpArrow /d " & Chr(34) & "%候选%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%" & Chr(34) & " /v Hand /d " & Chr(34) & "%连接选择%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 添加鼠标方案" & vbCrLf
BAT_TEXT = BAT_TEXT & "%Cursor%\Schemes" & Chr(34) & " /v %鼠标方案% /d " & Chr(34) & "%正常选择%,%帮助选择%,%后台运行%,%忙%,%精确定位%,%选定文本%,%手写%,%不可用%,%垂直调整%,%水平调整%,%对角线1%,%对角线2%,%移动%,%候选%,%连接选择%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf

'声音
    BAT_TEXT = BAT_TEXT & "REM ★改变声音" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM 下面的su是方案路径，建议英文数字（相当于方案的英文名称），方案名字可以设置中文" & vbCrLf
If Main.Check_snd.value = 0 Or Main.Combo_Sys_Snd.ListIndex = 0 Then
    If Main.sound_name_C.text <> "" And Main.sound_name_E.text <> "" Then
        BAT_TEXT = BAT_TEXT & "set su=" & Main.sound_name_E.text & vbCrLf
        BAT_TEXT = BAT_TEXT & "set 方案名字=" & Main.sound_name_C.text & vbCrLf
    ElseIf Main.sound_name_C.text = "" And Main.sound_name_E.text <> "" Then
        BAT_TEXT = BAT_TEXT & "set su=" & Main.sound_name_E.text & vbCrLf
        BAT_TEXT = BAT_TEXT & "set 方案名字=" & Main.sound_name_E.text & vbCrLf
    ElseIf Main.sound_name_C.text <> "" And Main.sound_name_E.text = "" Then
        BAT_TEXT = BAT_TEXT & "set su=" & Main.sound_name_C.text & vbCrLf
        BAT_TEXT = BAT_TEXT & "set 方案名字=" & Main.sound_name_C.text & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "set su=none" & vbCrLf
        BAT_TEXT = BAT_TEXT & "set 方案名字=none" & vbCrLf
    End If
Else
    BAT_TEXT = BAT_TEXT & "set su=" & Sound_Name(Main.Combo_Sys_Snd.ListIndex) & vbCrLf
    BAT_TEXT = BAT_TEXT & "set 方案名字=" & GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(Main.Combo_Sys_Snd.ListIndex), vbNullString) & vbCrLf
End If

    BAT_TEXT = BAT_TEXT & "REM 设置路径" & vbCrLf
    For i = 0 To UBound(Sound, 2)
        BAT_TEXT = BAT_TEXT & "set " & Sound(0, i) & "=" & Sound(2, i) & vbCrLf
    Next
    BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set sound=reg add " & Chr(34) & "HKEY_CURRENT_USER\AppEvents\Schemes\Apps\" & vbCrLf
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
    '设置当前
    n = 0
    For i = 0 To UBound(F_Sound)
    BAT_TEXT = BAT_TEXT & "REM 设置当前，" & F_Sound(i) & vbCrLf
        For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
                BAT_TEXT = BAT_TEXT & "%sound%" & F_Sound(i) & "\" & Sound(0, n) & "\.Current" & Chr(34) & " /f /ve /d " & Chr(34) & "%" & Sound(0, n) & "%" & Chr(34) & vbCrLf
            n = n + 1
        Next j
    Next i
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
'添加方案
    BAT_TEXT = BAT_TEXT & "Reg Add " & Chr(34) & "HKEY_CURRENT_USER\AppEvents\Schemes\Names\%su%" & Chr(34) & " /f /ve /d " & Chr(34) & "%方案名字%" & Chr(34) & vbCrLf
    BAT_TEXT = BAT_TEXT & "REM 选择为当前方案" & vbCrLf
    BAT_TEXT = BAT_TEXT & "Reg Add HKCU\AppEvents\Schemes /f /ve /d %su%" & vbCrLf
    '方案
    n = 0
    For i = 0 To UBound(F_Sound)
    BAT_TEXT = BAT_TEXT & "REM 设置方案，" & F_Sound(i) & vbCrLf
        For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
            If Sound(2, n) <> "" Then
                BAT_TEXT = BAT_TEXT & "%sound%" & F_Sound(i) & "\" & Sound(0, n) & "\%su%" & Chr(34) & " /f /ve /d " & Chr(34) & "%" & Sound(0, n) & "%" & Chr(34) & vbCrLf
            End If
            n = n + 1
        Next j
    Next i

BAT_TEXT = BAT_TEXT & "" & vbCrLf


BAT_TEXT = BAT_TEXT & "REM ★改变主题风格" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM ◢↓◣" & vbCrLf
If Main.url_mss.text <> "" And Main.mss_Classic.value = False Then
    BAT_TEXT = BAT_TEXT & "set 视觉文件地址=" & Main.url_mss.text & vbCrLf
ElseIf Main.mss_Classic.value = True Then
    BAT_TEXT = BAT_TEXT & "set 视觉文件地址=" & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "set 视觉文件地址=%SystemRoot%\resources\Themes\Aero\Aero.msstyles" & vbCrLf
End If
BAT_TEXT = BAT_TEXT & "REM 以下几个颜色可以在个性化里面调节或者魔方调节(Aero效果调节)后保存，然后使用附带的“读取视觉风格设置值.bat”查看" & vbCrLf
BAT_TEXT = BAT_TEXT & "set 主颜色=" & Main.Value_ColorizationColor.text & vbCrLf
BAT_TEXT = BAT_TEXT & "set 主颜色平衡=" & Main.Value_ColorizationColorBalance.text & vbCrLf
BAT_TEXT = BAT_TEXT & "set 发光颜色=" & Main.Value_ColorizationAfterglow.text & vbCrLf
BAT_TEXT = BAT_TEXT & "set 发光颜色平衡=" & Main.Value_ColorizationAfterglowBalance.text & vbCrLf
BAT_TEXT = BAT_TEXT & "set 模糊平衡=" & Main.Value_ColorizationBlurBalance.text & vbCrLf
BAT_TEXT = BAT_TEXT & "set 大背景透明度=" & Main.Value_ColorizationGlassReflectionIntensity.text & vbCrLf
If Main.mss_Aero.value = True Then
    BAT_TEXT = BAT_TEXT & "set 开Aero=1" & vbCrLf
ElseIf Main.mss_Basic.value = True Then
    BAT_TEXT = BAT_TEXT & "set 开Aero=0" & vbCrLf
End If
If Main.Check_Alpha.value = 0 Then
    BAT_TEXT = BAT_TEXT & "set 透明=1" & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "set 透明=0" & vbCrLf
End If
BAT_TEXT = BAT_TEXT & "REM ◥↑◤" & vbCrLf
BAT_TEXT = BAT_TEXT & "set style=reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Microsoft\Windows" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\CurrentVersion\ThemeManager" & Chr(34) & " /v DllName /d " & Chr(34) & "%视觉文件地址%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 打开视觉风格" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\CurrentVersion\ThemeManager" & Chr(34) & " /v ThemeActive /d " & Chr(34) & "1" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 变为Aero，要Basic则为0" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v Composition /t REG_DWORD /d " & Chr(34) & "%开Aero%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 开启透明,0为透明，1为不透明" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationOpaqueBlend /t REG_DWORD /d " & Chr(34) & "%透明%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 主颜色" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationColor /t REG_DWORD /d " & Chr(34) & "%主颜色%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 主颜色平衡（深浅，0~100）" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationColorBalance /t REG_DWORD /d " & Chr(34) & "%主颜色平衡%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 发光颜色（对于没开透明的没用）" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationAfterglow /t REG_DWORD /d " & Chr(34) & "%发光颜色%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 发光颜色平衡（同上,0~100）" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationAfterglowBalance /t REG_DWORD /d " & Chr(34) & "%发光颜色平衡%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 模糊平衡（模糊量,0~100）" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationBlurBalance /t REG_DWORD /d " & Chr(34) & "%模糊平衡%" & Chr(34) & " /f" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 大背景透明度（AERO条纹数量,0~100）" & vbCrLf
BAT_TEXT = BAT_TEXT & "%style%\DWM" & Chr(34) & " /v ColorizationGlassReflectionIntensity /t REG_DWORD /d " & Chr(34) & "%大背景透明度%" & Chr(34) & " /f" & vbCrLf

If Main.mss_Classic.value = True Then
    BAT_TEXT = BAT_TEXT & "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ThemeManager\ThemeActive" & Chr(34) & " /v ThemeActive /d " & Chr(34) & "0" & Chr(34) & " /f" & vbCrLf
Else

    BAT_TEXT = BAT_TEXT & "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ThemeManager\ThemeActive" & Chr(34) & " /v ThemeActive /d " & Chr(34) & "1" & Chr(34) & " /f" & vbCrLf
End If

If Main.Check_insert_system_color.value <> 0 Or Main.mss_Classic.value = True Then
    BAT_TEXT = BAT_TEXT & "Rem ★改变系统颜色" & vbCrLf
    BAT_TEXT = BAT_TEXT & "Rem ◢↓◣" & vbCrLf
    For i = 0 To 30
        BAT_TEXT = BAT_TEXT & "Set" & SysColors(i, 0) & " =" & SysColors(i, Main.ImageCombo_Classic_Style.SelectedItem.Index) & vbCrLf
    Next
    BAT_TEXT = BAT_TEXT & "Rem ◥↑◤" & vbCrLf
    BAT_TEXT = BAT_TEXT & "set color=reg add " & Chr(34) & "HKEY_CURRENT_USER\Control Panel\Colors" & vbCrLf
    For i = 0 To 30
        BAT_TEXT = BAT_TEXT & "%color%" & Chr(34) & " /v " & SysColors(i, 0) & " /d " & Chr(34) & Chr(37) & SysColors(i, 0) & Chr(37) & Chr(34) & " /f" & vbCrLf
    Next
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
End If

BAT_TEXT = BAT_TEXT & "" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 调用动态链接库立即刷新桌面壁纸" & vbCrLf
BAT_TEXT = BAT_TEXT & "RunDll32.exe USER32.DLL,UpdatePerUserSystemParameters" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 重启主题服务以立即启用主题" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("正在重启主题服务以立即使用主题风格", "Save_Theme", "Restart_Theme_Service", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "net stop Themes" & vbCrLf
BAT_TEXT = BAT_TEXT & "net start Themes" & vbCrLf
BAT_TEXT = BAT_TEXT & "REM 打开鼠标更换" & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("请在打开的窗口点击确定更换鼠标样式", "Save_Theme", "Change_Cursors", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "rundll32.exe shell32.dll,Control_RunDLL main.cpl @0,1" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf

If Sound(0, 2) <> "" And Main.Check_snd.value <> 1 Then '为空则不播放
    BAT_TEXT = BAT_TEXT & "REM 播放音乐" & vbCrLf
    BAT_TEXT = BAT_TEXT & "echo strSoundFile = " & Chr(34) & "%ChangeTheme%" & Chr(34) & ">>%temp%\" & Main.T_name_C.text & "sound.vbs" & vbCrLf
    BAT_TEXT = BAT_TEXT & "echo Set objShell = CreateObject(" & Chr(34) & "Wscript.Shell" & Chr(34) & ")>>%temp%\" & Main.T_name_C.text & "sound.vbs" & vbCrLf
    BAT_TEXT = BAT_TEXT & "echo strCommand = " & Chr(34) & "wmplayer /play /close " & Chr(34) & " ^& chr(34) ^& strSoundFile ^& chr(34)>>%temp%\" & Main.T_name_C.text & "sound.vbs" & vbCrLf
    BAT_TEXT = BAT_TEXT & "echo objShell.Run strCommand, 0, True>>%temp%\" & Main.T_name_C.text & "sound.vbs" & vbCrLf
    BAT_TEXT = BAT_TEXT & "cscript //nologo %temp%\" & Main.T_name_C.text & "sound.vbs & del %temp%\" & Main.T_name_C.text & "sound.vbs" & vbCrLf
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
End If

BAT_TEXT = BAT_TEXT & "cls" & vbCrLf
BAT_TEXT = BAT_TEXT & "title " & Load_Lanuage("应用完成", "Save_Theme", "Title_Over", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("应用主题完成", "Save_Theme", "Aply_Theme_Done", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo " & Load_Lanuage("图标、壁纸等可能无法立即刷新显示，注销再登入就能全部显示了", "Save_Theme", "Refresh", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "echo." & vbCrLf
BAT_TEXT = BAT_TEXT & "pause" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf

BAT_TEXT = BAT_TEXT & ":exit" & vbCrLf
BAT_TEXT = BAT_TEXT & "exit" & vbCrLf

Print #1, BAT_TEXT
Close #1
End Sub

Public Sub Save_Theme(ByVal save_url As String, Optional ByVal System_Ver As Byte = 0, Optional ByVal Win8 As Boolean = False)
Open save_url For Output As #1

Dim i%, j%, n%
Dim BAT_TEXT As String
Dim text_num_temp As Single
Dim text_temp1, text_temp2 As String

BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("此主题文件由枫谷主题", "Save_Theme", "Theme_From1", Lanuage_Now) & " V" & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision & Load_Lanuage("生成", "Save_Theme", "Theme_From2", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("作者：", "Save_Theme", "Author", Lanuage_Now) & Main.Maker_Name.text & vbCrLf
BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("网址：", "Save_Theme", "Web", Lanuage_Now) & Main.Maker_Web_Url.text & vbCrLf
If Main.Maker_Introduce.text <> "" Then
BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("版权信息：", "Save_Theme", "Copyright", Lanuage_Now) & vbCrLf '把换行符替换为一个换行符+;
BAT_TEXT = BAT_TEXT & ";" & Replace(Main.Maker_Introduce, vbCrLf, vbCrLf & ";") & vbCrLf '把换行符替换为一个换行符+;
End If
BAT_TEXT = BAT_TEXT & "" & vbCrLf
'主题信息
BAT_TEXT = BAT_TEXT & "[Theme]" & vbCrLf
BAT_TEXT = BAT_TEXT & "DisplayName=" & Main.T_name_C.text & vbCrLf
BAT_TEXT = BAT_TEXT & "BrandImage=" & Main.url_Tlogo.text & vbCrLf
BAT_TEXT = BAT_TEXT & "SetLogonBackground=1" & vbCrLf
'图标

BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("我的电脑", "Main", "Icon_Name0", Lanuage_Now) & vbCrLf
BAT_TEXT = BAT_TEXT & "[CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon]" & vbCrLf
If Main.url_icon(0).text <> "" Then
    BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(0, 0) & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(0, 3) & vbCrLf
End If

BAT_TEXT = BAT_TEXT & "[CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon]" & vbCrLf
BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("回收站（空）", "Main", "Icon_Name3", Lanuage_Now) & vbCrLf
If Main.url_icon(3).text <> "" Then
    BAT_TEXT = BAT_TEXT & "Empty=" & SysIco(3, 0) & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "Empty=" & SysIco(3, 3) & vbCrLf
End If
BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("回收站（满）", "Main", "Icon_Name4", Lanuage_Now) & vbCrLf
If Main.url_icon(4).text <> "" Then
    BAT_TEXT = BAT_TEXT & "Full=" & SysIco(4, 0) & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "Full=" & SysIco(4, 3) & vbCrLf
End If

If System_Ver >= 3 Or System_Ver = 0 Then 'Win7以上
    BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("我的文档", "Main", "Icon_Name1", Lanuage_Now) & vbCrLf
    BAT_TEXT = BAT_TEXT & "[CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon]" & vbCrLf
    If Main.url_icon(1).text <> "" Then
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(1, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(1, 3) & vbCrLf
    End If
    
    BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("网上邻居", "Main", "Icon_Name2", Lanuage_Now) & vbCrLf
    BAT_TEXT = BAT_TEXT & "[CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon]" & vbCrLf
    If Main.url_icon(2).text <> "" Then
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(2, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(2, 3) & vbCrLf
    End If
End If

If System_Ver <= 2 Then 'Vista以下
    BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("我的文档", "Main", "Icon_Name1", Lanuage_Now) & vbCrLf
    BAT_TEXT = BAT_TEXT & "[CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon]" & vbCrLf
    If Main.url_icon(1).text <> "" Then
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(1, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(1, 3) & vbCrLf
    End If
    
    BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("网上邻居", "Main", "Icon_Name2", Lanuage_Now) & vbCrLf
    BAT_TEXT = BAT_TEXT & "[CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon]" & vbCrLf
    If Main.url_icon(2).text <> "" Then
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(2, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(2, 3) & vbCrLf
    End If
End If

If System_Ver = 0 Then
    BAT_TEXT = BAT_TEXT & ";" & Load_Lanuage("Internet Explorer", "Main", "Icon_Name5", Lanuage_Now) & vbCrLf
    BAT_TEXT = BAT_TEXT & "[CLSID\{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon]" & vbCrLf
    If Main.url_icon(2).text <> "" Then
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(5, 0) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "DefaultValue=" & SysIco(5, 3) & vbCrLf
    End If
End If
BAT_TEXT = BAT_TEXT & "" & vbCrLf

'鼠标
BAT_TEXT = BAT_TEXT & "[Control Panel\Cursors]" & vbCrLf
For i = 0 To 14
    BAT_TEXT = BAT_TEXT & SysCur(i, 1) & "=" & SysCur(i, 0) & vbCrLf
Next
BAT_TEXT = BAT_TEXT & "DefaultValue=" & Main.name_cur.text & vbCrLf
BAT_TEXT = BAT_TEXT & "DefaultValue.MUI=" & Main.name_cur.text & vbCrLf

BAT_TEXT = BAT_TEXT & "" & vbCrLf
'壁纸
BAT_TEXT = BAT_TEXT & "[Control Panel\Desktop]" & vbCrLf
BAT_TEXT = BAT_TEXT & "Wallpaper=" & Main.url_paper.text & vbCrLf
BAT_TEXT = BAT_TEXT & "TileWallpaper=" & TileWallpaper_value & vbCrLf
If System_Ver <= 2 And WallpaperStyle_value >= 6 Then '只有7以上才支持壁纸填充
    BAT_TEXT = BAT_TEXT & "WallpaperStyle=2" & vbCrLf
Else
    BAT_TEXT = BAT_TEXT & "WallpaperStyle=" & WallpaperStyle_value & vbCrLf
End If
BAT_TEXT = BAT_TEXT & "Pattern=" & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf


If System_Ver >= 3 Or System_Ver = 0 Then '7以上
    '壁纸幻灯片
    If Main.url_paper_files.text <> "" Then
        BAT_TEXT = BAT_TEXT & "[Slideshow]" & vbCrLf
        BAT_TEXT = BAT_TEXT & "ImagesRootPath=" & Main.url_paper_files.text & vbCrLf
        If Main.Combo_Paper_Change_Time.ListIndex = 0 Then    '10秒
            BAT_TEXT = BAT_TEXT & "Interval=10000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 1 Then  '30秒
            BAT_TEXT = BAT_TEXT & "Interval=30000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 2 Then   '1分钟
            BAT_TEXT = BAT_TEXT & "Interval=60000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 3 Then   '3分钟
            BAT_TEXT = BAT_TEXT & "Interval=180000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 4 Then    '5分钟
            BAT_TEXT = BAT_TEXT & "Interval=300000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 5 Then    '10分钟
            BAT_TEXT = BAT_TEXT & "Interval=600000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 6 Then   '15分钟
            BAT_TEXT = BAT_TEXT & "Interval=900000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 7 Then   '20分钟
            BAT_TEXT = BAT_TEXT & "Interval=1200000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 8 Then   '30分钟
            BAT_TEXT = BAT_TEXT & "Interval=1800000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 9 Then  '1小时
            BAT_TEXT = BAT_TEXT & "Interval=3600000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 10 Then   '2小时
            BAT_TEXT = BAT_TEXT & "Interval=7200000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 11 Then  '3小时
            BAT_TEXT = BAT_TEXT & "Interval=10800000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 12 Then   '6小时
            BAT_TEXT = BAT_TEXT & "Interval=21600000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 13 Then   '12小时
            BAT_TEXT = BAT_TEXT & "Interval=43200000" & vbCrLf
        ElseIf Main.Combo_Paper_Change_Time.ListIndex = 14 Then     '一天
            BAT_TEXT = BAT_TEXT & "Interval=84600000" & vbCrLf
        Else
            BAT_TEXT = BAT_TEXT & "Interval=1800000" & vbCrLf '30分钟
        End If
        If Main.Check_paper_change.value <> 0 Then
            BAT_TEXT = BAT_TEXT & "Shuffle=1" & vbCrLf
        Else
            BAT_TEXT = BAT_TEXT & "Shuffle=0" & vbCrLf
        End If
        
        n = 0
        If Main.Papers_Edit_Allow.value <> 0 And Main.TreeView_paper.Nodes.count > 0 Then
            'For i = 1 To UBound(PaperFileName)
            For i = 1 To PaperFileName.count
                If Main.TreeView_paper.Nodes(i).Checked = True Then
                    BAT_TEXT = BAT_TEXT & "Item" & n & "Path=" & url_to_P(PaperFileName(i)) & vbCrLf
                    n = n + 1
                End If
            Next
        End If
        
        BAT_TEXT = BAT_TEXT & "" & vbCrLf
        
        n = 0
        If Main.Papers_Edit_Allow.value <> 0 And Main.TreeView_paper.Nodes.count > 0 Then
            BAT_TEXT = BAT_TEXT & "[Slideshow1]" & vbCrLf
            'For i = 1 To UBound(PaperFileName)
            For i = 1 To PaperFileName.count
                If Main.TreeView_paper.Nodes(i).Checked = True Then
                    BAT_TEXT = BAT_TEXT & "Item" & n & "Path=" & url_to_P(PaperFileName(i)) & vbCrLf
                    n = n + 1
                End If
            Next
        End If
    End If
End If
'屏幕保护程序
BAT_TEXT = BAT_TEXT & "[boot]" & vbCrLf
BAT_TEXT = BAT_TEXT & "SCRNSAVE.EXE=" & Main.url_scr.text & vbCrLf
BAT_TEXT = BAT_TEXT & "" & vbCrLf

'声音
BAT_TEXT = BAT_TEXT & "[Sounds]" & vbCrLf
n = 0
If Main.Check_snd.value = 0 Or Main.Combo_Sys_Snd.ListIndex = 0 Then '没选中“使用系统声音”或者选中的是当前音效
    BAT_TEXT = BAT_TEXT & "SchemeName=" & Main.sound_name_C.text & vbCrLf
    BAT_TEXT = BAT_TEXT & "" & vbCrLf
    For i = 0 To UBound(F_Sound)
        For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
            If Sound(2, n) <> "" Then
                BAT_TEXT = BAT_TEXT & "[AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n) & "]" & vbCrLf
                BAT_TEXT = BAT_TEXT & "DefaultValue=" & Sound(2, n) & vbCrLf
            End If
            n = n + 1
        Next j
    Next i
Else
    If Left$(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(Main.Combo_Sys_Snd.ListIndex), vbNullString), 1) = "@" Then
        BAT_TEXT = BAT_TEXT & "SchemeName=" & GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(Main.Combo_Sys_Snd.ListIndex), vbNullString) & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "SchemeName=" & GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(Main.Combo_Sys_Snd.ListIndex), vbNullString) & vbCrLf
        BAT_TEXT = BAT_TEXT & "" & vbCrLf
        For i = 0 To UBound(F_Sound)
            For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
                If Sound(2, n) <> "" Then
                    BAT_TEXT = BAT_TEXT & "[AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n) & "]" & vbCrLf
                    BAT_TEXT = BAT_TEXT & "DefaultValue=" & Sound(2, n) & vbCrLf
                End If
                n = n + 1
            Next j
        Next i
    End If
End If
BAT_TEXT = BAT_TEXT & "" & vbCrLf

'视觉风格
BAT_TEXT = BAT_TEXT & "[VisualStyles]" & vbCrLf
If Main.mss_Classic.value = True Then
    BAT_TEXT = BAT_TEXT & "Path=" & vbCrLf
    If Main.ImageCombo_Classic_Style.SelectedItem.Index = 1 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-854" & vbCrLf
    ElseIf Main.ImageCombo_Classic_Style.SelectedItem.Index = 2 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-854" & vbCrLf
    ElseIf Main.ImageCombo_Classic_Style.SelectedItem.Index = 3 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-850" & vbCrLf
    ElseIf Main.ImageCombo_Classic_Style.SelectedItem.Index = 4 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-851 & vbCrLf"
    ElseIf Main.ImageCombo_Classic_Style.SelectedItem.Index = 5 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-852" & vbCrLf
    ElseIf Main.ImageCombo_Classic_Style.SelectedItem.Index = 6 Then
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-853" & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-854" & vbCrLf
    End If
    BAT_TEXT = BAT_TEXT & "Size=@themeui.dll,-2019" & vbCrLf
Else
    If Main.url_mss.text <> "" And Win8 = False Then
        BAT_TEXT = BAT_TEXT & "Path=" & Main.url_mss.text & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "Path=%ResourceDir%\Themes\Aero\Aero.msstyles" & vbCrLf
    End If
    BAT_TEXT = BAT_TEXT & "ColorStyle=@themeui.dll,-2027" & vbCrLf
    BAT_TEXT = BAT_TEXT & "Size=@themeui.dll,-2028" & vbCrLf
End If

If System_Ver >= 2 Or System_Ver = 0 Then 'Vista以上
    BAT_TEXT = BAT_TEXT & "ColorizationColor=" & Main.Value_ColorizationColor.text & vbCrLf
    '透明
    If Main.Check_Alpha.value = 0 Then
        BAT_TEXT = BAT_TEXT & "Transparency=0" & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "Transparency=1" & vbCrLf
    End If
    BAT_TEXT = BAT_TEXT & "VisualStyleVersion=10" & vbCrLf
    'Aero/Basic
    If Main.mss_Aero.value = True Then
        BAT_TEXT = BAT_TEXT & "Composition=1" & vbCrLf
    ElseIf Main.mss_Basic.value = True Then
        BAT_TEXT = BAT_TEXT & "Composition=0" & vbCrLf
    End If
End If
If Main.mss_Classic.value = True Or Main.Check_insert_system_color.value <> 0 Then
    BAT_TEXT = BAT_TEXT & "[Metrics]" & vbCrLf
    If Main.ImageCombo_Classic_Style.SelectedItem.Index = 1 Then
        BAT_TEXT = BAT_TEXT & "NonclientMetrics=88 1 0 0 1 0 0 0 16 0 0 0 16 0 0 0 18 0 0 0 18 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 188 2 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 12 0 0 0 15 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 188 2 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 18 0 0 0 18 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 144 1 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 144 1 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 144 1 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 " & vbCrLf
        BAT_TEXT = BAT_TEXT & "LangID=2052" & vbCrLf
        BAT_TEXT = BAT_TEXT & "IconMetrics=76 0 0 0 75 0 0 0 74 0 0 0 1 0 0 0 244 255 255 255 0 0 0 0 0 0 0 0 0 0 0 0 144 1 0 0 0 0 0 1 0 0 0 0 203 206 204 229 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 " & vbCrLf
    Else
        BAT_TEXT = BAT_TEXT & "CaptionFont=@themeui.dll,-2037" & vbCrLf
        BAT_TEXT = BAT_TEXT & "SmCaptionFont=@themeui.dll,-2038" & vbCrLf
        BAT_TEXT = BAT_TEXT & "MenuFont=@themeui.dll,-2039" & vbCrLf
        BAT_TEXT = BAT_TEXT & "StatusFont=@themeui.dll,-2040" & vbCrLf
        BAT_TEXT = BAT_TEXT & "MessageFont=@themeui.dll,-2041" & vbCrLf
        BAT_TEXT = BAT_TEXT & "IconFont=@themeui.dll,-2042" & vbCrLf
    End If
    If Main.Check_insert_system_color.value <> 0 And Main.ImageCombo_Classic_Style.SelectedItem.Index = 1 Then
        BAT_TEXT = BAT_TEXT & "[Control Panel\Colors]" & vbCrLf
        For i = 0 To 30
            BAT_TEXT = BAT_TEXT & SysColors(i, 0) & "=" & SysColors(i, Main.ImageCombo_Classic_Style.SelectedItem.Index) & vbCrLf
        Next
        Main.ImageCombo_Classic_Style.ComboItems(1).Selected = True
    End If
End If
BAT_TEXT = BAT_TEXT & "" & vbCrLf

'主题必有的信息，不同系统版本
BAT_TEXT = BAT_TEXT & "[MasterThemeSelector]" & vbCrLf
If System_Ver = 4 Then
    BAT_TEXT = BAT_TEXT & "MTSM=RJSPBS" & vbCrLf 'win8则是RJSPBS
Else
    BAT_TEXT = BAT_TEXT & "MTSM=DABJDKT" & vbCrLf
End If
BAT_TEXT = BAT_TEXT & "" & vbCrLf

Print #1, BAT_TEXT
Close #1
End Sub
'直接应用主题
Public Sub Aply_Theme(Load_Url As String)
Dim i%, j%, n%
Dim Slideshow_Url As String
Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Themes", "CurrentTheme", Load_Url)
'壁纸
If Main.url_paper <> "" Then '为空则无壁纸
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", Main.url_paper)
    If System_Ver <= 2 And WallpaperStyle_value >= 6 Then '只有7以上才支持壁纸填充
        Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", 2)
    Else
        Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", CStr(WallpaperStyle_value))
    End If
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", CStr(TileWallpaper_value))
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\General", "WallpaperSource", Main.url_paper)
    Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, url_to_S(Main.url_paper), 0)
End If

'传递给自动更换壁纸
If Main.url_paper_files <> "" Then '为空则无壁纸文件夹
    Call SavePaperList(New_List)
End If
 
'屏保
If Main.url_scr <> "" Then '为空则无屏保
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", Main.url_scr)
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveTimeOut", Main.scr_wait_min * 60)
Else
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", "")
End If

'图标
If SysIco(0, 0) <> "" Then '我的电脑
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon", vbNullString, SysIco(0, 0))
Else
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon", vbNullString, SysIco(0, 3))
End If
If System_Ver < 6.1 Then
    If SysIco(0, 0) <> "" Then '我的文档
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", vbNullString, SysIco(1, 0))
    Else
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", vbNullString, SysIco(1, 3))
    End If
    If SysIco(0, 0) <> "" Then '网上邻居
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", vbNullString, SysIco(2, 0))
    Else
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", vbNullString, SysIco(2, 3))
    End If
Else
    If SysIco(0, 0) <> "" Then '我的文档
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon", vbNullString, SysIco(1, 0))
    Else
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon", vbNullString, SysIco(1, 3))
    End If
    If SysIco(0, 0) <> "" Then '网上邻居
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon", vbNullString, SysIco(2, 0))
    Else
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon", vbNullString, SysIco(2, 3))
    End If
End If
If SysIco(0, 0) <> "" Then '回收站空
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Empty", SysIco(3, 0))
Else
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Empty", SysIco(3, 3))
End If
If SysIco(0, 0) <> "" Then '回收站满
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Full", SysIco(4, 0))
Else
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Full", SysIco(4, 3))
End If
If SysIco(0, 0) <> "" Then 'IE
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon", vbNullString, SysIco(5, 0))
Else
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon", vbNullString, SysIco(5, 3))
End If
'鼠标
Dim Cur_Temp As String
For i = 0 To 14
    Call SetString(HKEY_CURRENT_USER, "Control Panel\Cursors", SysCur(i, 1), SysCur(i, 0))
    Cur_Temp = Cur_Temp & SysCur(i, 0) & ","
Next
Cur_Temp = Mid(Cur_Temp, 1, Len(Cur_Temp) - 1)
Call SetString(HKEY_CURRENT_USER, "Control Panel\Cursors\Schemes", Main.name_cur, Cur_Temp)


'声音
Dim su As String, 方案名字 As String
If Main.Check_snd.value = 0 Or Main.Combo_Sys_Snd.ListIndex = 0 Then
    If Main.sound_name_C.text <> "" And Main.sound_name_E.text <> "" Then
        su = Main.sound_name_E
        方案名字 = Main.sound_name_C
    ElseIf Main.sound_name_C.text = "" And Main.sound_name_E.text <> "" Then
        su = Main.sound_name_E
        方案名字 = Main.sound_name_E
    ElseIf Main.sound_name_C.text <> "" And Main.sound_name_E.text = "" Then
        su = Main.sound_name_C
        方案名字 = Main.sound_name_C
    Else
        su = "none"
        方案名字 = "none"
    End If
    
    '添加方案
    Call SetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & su, vbNullString, 方案名字)
    Call SetString(HKEY_CURRENT_USER, "AppEvents\Schemes", vbNullString, su)

    '方案
    n = 0
    For i = 0 To UBound(F_Sound)
        For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
            If Sound(2, n) <> "" Then
                Call SetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n) & "\" & su, vbNullString, url_to_N(Sound(2, n)))
            
            End If
            n = n + 1
        Next j
    Next i
    
Else
    su = Sound_Name(Main.Combo_Sys_Snd.ListIndex)
    方案名字 = GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(Main.Combo_Sys_Snd.ListIndex), vbNullString)
End If

    '设置当前
    n = 0
    For i = 0 To UBound(F_Sound)
        For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
                Call SetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n) & "\.Current", vbNullString, url_to_N(Sound(2, n)))
            n = n + 1
        Next j
    Next i

Dim 视觉文件地址 As String, 主颜色 As Long, 发光颜色 As Long
Dim 主颜色平衡 As Byte, 发光颜色平衡 As Byte, 模糊平衡 As Byte, 大背景透明度 As Byte
Dim 开Aero As Byte, 透明 As Byte
If Main.url_mss.text <> "" And Main.mss_Classic.value = False Then
    视觉文件地址 = url_to_S(Main.url_mss)
ElseIf Main.mss_Classic.value = True Then
    视觉文件地址 = ""
Else
    视觉文件地址 = "%SystemRoot%\resources\Themes\Aero\Aero.msstyles"
End If
主颜色 = x16_to_x10(Mid(Main.Value_ColorizationColor.text, 3))
主颜色平衡 = Main.HScroll_ColorizationColorBalance.value
发光颜色 = x16_to_x10(Mid(Main.Value_ColorizationAfterglow.text, 3))
发光颜色平衡 = Main.HScroll_ColorizationAfterglowBalance.value
模糊平衡 = Main.HScroll_ColorizationBlurBalance.value
大背景透明度 = Main.HScroll_ColorizationGlassReflectionIntensity.value
If Main.mss_Aero.value = True Then
    开Aero = 1
ElseIf Main.mss_Basic.value = True Then
    开Aero = 0
End If
If Main.Check_Alpha.value = 0 Then
    透明 = 1
Else
    透明 = 0
End If

Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", 视觉文件地址)

Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "Composition", 开Aero)

Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationOpaqueBlend", 透明)

Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColor", 主颜色)
Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColorBalance", 主颜色平衡)
Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglow", 发光颜色)
Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglowBalance", 发光颜色平衡)
Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationBlurBalance", 模糊平衡)
Call SetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationGlassReflectionIntensity", 大背景透明度)

If Main.mss_Classic.value = True Then 'classic
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "0")
Else
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
End If

If Main.Check_insert_system_color.value <> 0 Or Main.mss_Classic.value = True Then
    For i = 0 To 30
        Call SetString(HKEY_CURRENT_USER, "Control Panel\Colors", SysColors(i, 0), SysColors(i, Main.ImageCombo_Classic_Style.SelectedItem.Index))
    Next
End If


    Call NeiCun_Timer '清理内存，不然就一个劲疯涨
    
End Sub
'保存壁纸列表到自动壁纸目录
Public Sub SavePaperList(New_List As String)
    Dim x As Integer
    If AutoPaper > 0 Then
        If AutoPaper = 2 Then  ' 如果用户单击No按钮，则停止Unload事件。
            x = MsgBox(Load_Lanuage("是否将壁纸列表传送给“自动更换壁纸”程序？", "Main", "SendPaperList", Lanuage_Now), 36)
        Else
            x = vbYes
        End If
        
        If x = vbYes Then
        '传递壁纸列表
            Dim PaperList_Url As String
            PaperList_Url = Environ("AppData") & "\MapleAutoWallpaper\PaperList.txt"
            Open PaperList_Url For Output As #1
            Print #1, New_List
            Close #1
            
            Dim MapleAutoWallpaper_Config_Url As String
            MapleAutoWallpaper_Config_Url = Environ("AppData") & "\MapleAutoWallpaper\config.ini"
        '传递时间
            With Main.Combo_Paper_Change_Time
                If .ListIndex = 0 Then    '10秒
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 10000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 1 Then  '30秒
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 30000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 2 Then   '1分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 60000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 3 Then   '3分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 180000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 4 Then    '5分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 300000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 5 Then    '10分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 600000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 6 Then   '15分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 900000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 7 Then   '20分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 1200000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 8 Then   '30分钟
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 1800000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 9 Then  '1小时
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 3600000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 10 Then   '2小时
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 7200000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 11 Then  '3小时
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 10800000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 12 Then   '6小时
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 21600000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 13 Then   '12小时
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 43200000, MapleAutoWallpaper_Config_Url)
                ElseIf .ListIndex = 14 Then     '一天
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 84600000, MapleAutoWallpaper_Config_Url)
                Else
                    Call WriteIni("Aoto_Change_Wallpaper", "ChangeTime", 1800000, MapleAutoWallpaper_Config_Url) '30分钟
                End If
            End With
            Call WriteIni("Aoto_Change_Wallpaper", "A_New_List", 1, MapleAutoWallpaper_Config_Url)
        End If
    End If
End Sub
