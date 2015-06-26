Attribute VB_Name = "重画窗口语言"
Option Explicit
Public Function Load_Lanuage(ByVal Now_Show As String, ByVal strSectionHeader As String, ByVal strVariableName As String, Optional ByVal Change_Lanuage_Now As Integer = -1) As String
Dim text_temp() As String, text_temp2 As String
Dim i As Integer

If Change_Lanuage_Now = -1 Then
    Change_Lanuage_Now = Lanuage_Now
End If

    Erase text_temp '先清空数组
    text_temp2 = ""
    
    If Change_Lanuage_Now <> 0 Then
        If GetFromIni(strSectionHeader, strVariableName, Lanuages(Change_Lanuage_Now)) <> "" Then
            text_temp = Split(GetFromIni(strSectionHeader, strVariableName, Lanuages(Change_Lanuage_Now)), "|")
        Else
            text_temp = Split(Now_Show, "|")
        End If
    Else
        text_temp = Split(Now_Show, "|")
    End If
    
    For i = 0 To UBound(text_temp)
        If i = UBound(text_temp) Then
            text_temp2 = text_temp2 & text_temp(i)
        Else
            text_temp2 = text_temp2 & text_temp(i) & vbCrLf
        End If
    Next
    Load_Lanuage = text_temp2
End Function
Public Sub Change_Lanuage(ByVal Change_Lanuage_Now As String)
Dim text_temp() As String, text_temp2 As String
Dim i As Integer
Lanuage_Now = Change_Lanuage_Now
'启动引导窗口
With frmLoad
    .Caption = Load_Lanuage("枫谷主题 - 选择启动任务", "Load", "Caption", Lanuage_Now)
    .Frame_Basic.Caption = Load_Lanuage("Window7家庭普通版应用主题", "Load", "Frame_Basic", Lanuage_Now)
    .Frame_Edit.Caption = Load_Lanuage("编辑 / 生成Windows主题", "Load", "Frame_Edit", Lanuage_Now)
    .Command_Open_Control.Caption = Load_Lanuage("手动应用主题", "Load", "Command_Open_Control", Lanuage_Now)
    .Command_theme_to_Bat.Caption = Load_Lanuage("自动应用主题到系统", "Load", "Command_theme_to_Bat", Lanuage_Now)
    .Command_Edit.Caption = Load_Lanuage("打开编辑器", "Load", "Command_Edit", Lanuage_Now)
    .Check_frmLoad.Caption = Load_Lanuage("下次不再出现本界面", "Load", "Check_frmLoad", Lanuage_Now)
End With
'关于
With frmAbout
    .Caption = Load_Lanuage("关于", "About", "Caption", Lanuage_Now) & " " & Load_Lanuage("枫谷主题", "info", "AppName", Lanuage_Now)
    .lblVersion.Caption = Load_Lanuage("版本", "About", "Version", Lanuage_Now) & " " & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision
    .lblTitle.Caption = Load_Lanuage("枫谷主题", "info", "AppName", Lanuage_Now)
    .cmdOk.Caption = Load_Lanuage("确定", "About", "cmdOK", Lanuage_Now)
    .cmdVisitVeb.Caption = Load_Lanuage("访问官网", "About", "cmdVisitVeb", Lanuage_Now)
    .lblDescription.Caption = Load_Lanuage("本程序不是自动收集资源到主题文件夹生成主题，而是为已经安装的主题生成家庭版安装BAT|本程序适用于：|主题制作者生成让使用者能在Win7家庭普通版使用主题的BAT|使用主题者自行将没有添加Win7家庭版安装BAT的主题生成安装BAT", "About", "Description", Lanuage_Now)
    .lblDisclaimer.Caption = Load_Lanuage("本程序所有权及使用权归枫谷剑仙所有", "About", "Disclaimer", Lanuage_Now)
End With
'一键取色工具
With Get_color
    .Caption = Load_Lanuage("一键取色工具", "Get_color", "Caption", Lanuage_Now)
    .freshen.Caption = Load_Lanuage("刷新", "Get_color", "freshen", Lanuage_Now)
    .freshen2.Caption = Load_Lanuage("刷新", "Get_color", "freshen", Lanuage_Now)
    .add_all.Caption = Load_Lanuage("插入全部", "Get_color", "add_all", Lanuage_Now)
    .add_all2.Caption = Load_Lanuage("插入全部", "Get_color", "add_all", Lanuage_Now)
    .Command_mss.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationColor.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationColorBalance.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationAfterglow.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationAfterglowBalance.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationBlurBalance.Caption = Load_Lanuage("插入具", "Get_color", "add_one", Lanuage_Now)
    .Command_glass.Caption = Load_Lanuage("打开透明颜色面板", "Get_color", "Command_glass", Lanuage_Now)
    .Command_window.Caption = Load_Lanuage("打开窗体颜色和外观面板", "Get_color", "Command_window", Lanuage_Now)
    
    .Label_mss.Caption = Load_Lanuage("视觉风格文件", "Main", "Label_mss", Lanuage_Now)
    .Label_ColorizationColor.Caption = Load_Lanuage("主颜色", "Main", "Label_ColorizationColor", Lanuage_Now)
    .Label_ColorizationColorBalance.Caption = Load_Lanuage("主颜色平衡", "Main", "Label_ColorizationColorBalance", Lanuage_Now)
    .Label_ColorizationAfterglow.Caption = Load_Lanuage("发光颜色", "Main", "Label_ColorizationAfterglow", Lanuage_Now)
    .Label_ColorizationAfterglowBalance.Caption = Load_Lanuage("发光颜色平衡", "Main", "Label_ColorizationAfterglowBalance", Lanuage_Now)
    .Label_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("Aero条纹数量", "Main", "Label_ColorizationGlassReflectionIntensity", Lanuage_Now)
    .Label_ColorizationBlurBalance.Caption = Load_Lanuage("模糊平衡", "Main", "Label_ColorizationBlurBalance", Lanuage_Now)

    .Label_help.Caption = Load_Lanuage("使用方法：|先确保是在Aero风格下（Win7 HomeBasic下叫做Windows 7 Standard），使用Windows自带的个性化或者魔方（Aero效果调节）调节好您所满意的颜色，并保存|然后点击本工具的“刷新”，就会去读您当前的设置值，然后根据需要选择插入到主程序窗口里面去。|右边则是系统基本颜色设置，非调节Classic不用使用。", "Get_color", "help", Lanuage_Now)
End With
'导出向导
With CreatGuide
    .Caption = Load_Lanuage("主题生成向导", "CreatGuide", "Caption", Lanuage_Now)
    .cmdLast.Caption = Load_Lanuage("上一步", "CreatGuide", "cmdLast", Lanuage_Now)
    .cmdNext.Caption = Load_Lanuage("下一步", "CreatGuide", "cmdNext", Lanuage_Now)
    .cmdOk.Caption = Load_Lanuage("完成", "CreatGuide", "cmdOk", Lanuage_Now)
    .cmdCancel.Caption = Load_Lanuage("取消", "CreatGuide", "cmdCancel", Lanuage_Now)
    
    .Frame_Files.Caption = Load_Lanuage("请选择生成何种文件", "CreatGuide", "Files", Lanuage_Now)
    .Option_File(0).Caption = Load_Lanuage("Windows Theme文件", "CreatGuide", "Option_File_Theme", Lanuage_Now)
    .Option_File(1).Caption = Load_Lanuage("Bat文件（用于Win7家庭普通版）", "CreatGuide", "Option_File_Bat", Lanuage_Now)
    
    .Frame_Theme_Ver.Caption = Load_Lanuage("请选择您需要生成的版本", "CreatGuide", "Frame_Theme_Ver", Lanuage_Now)
    .Option_Theme_Ver(0).Caption = Load_Lanuage("Windows通用（注：视觉风格文件不能通用）", "CreatGuide", "Option_Theme_Ver_All", Lanuage_Now)
    .Option_Theme_Ver(1).Caption = Load_Lanuage("Windows XP / 2003", "CreatGuide", "Option_Theme_Ver_XP", Lanuage_Now)
    .Option_Theme_Ver(2).Caption = Load_Lanuage("Windows Vista / 2008", "CreatGuide", "Option_Theme_Ver_Vista", Lanuage_Now)
    .Option_Theme_Ver(3).Caption = Load_Lanuage("Windows 7 / 2008 R2", "CreatGuide", "Option_Theme_Ver_7", Lanuage_Now)
    .Option_Theme_Ver(4).Caption = Load_Lanuage("Windows 8", "CreatGuide", "Option_Theme_Ver_8", Lanuage_Now)
    
    .Frame_BT_Color.Caption = Load_Lanuage("请选择生成的BAT文件的文字与背景色", "CreatGuide", "Frame_BT_Color", Lanuage_Now)
    .Frame_BT_Color_Fore.Caption = Load_Lanuage("前景色", "CreatGuide", "Frame_BT_Color_Fore", Lanuage_Now)
    .Frame_BT_Color_Back.Caption = Load_Lanuage("背景色", "CreatGuide", "Frame_BT_Color_Back", Lanuage_Now)
End With
'设置
With Options
    .Caption = Load_Lanuage("选项设置", "OptionsForm", "Caption", Lanuage_Now)
    .Label_Lanuage.Caption = Load_Lanuage("软件语言/Lanuages", "OptionsForm", "Label_Lanuage", Lanuage_Now)
    .Command_Find_Lanuages.Caption = Load_Lanuage("获取更多Find More", "OptionsForm", "Command_Find_Lanuages", Lanuage_Now)
    .Frame_SystemTextShow.Caption = Load_Lanuage("部分系统名称", "OptionsForm", "Frame_SystemTextShow", Lanuage_Now)
    .SystemTextShow_Sys.Caption = Load_Lanuage("从系统读取", "OptionsForm", "SystemTextShow_Sys", Lanuage_Now)
    .SystemTextShow_ini.Caption = Load_Lanuage("从语言文件读取", "OptionsForm", "SystemTextShow_ini", Lanuage_Now)
    .Label_Snd_Style.Caption = Load_Lanuage("音效列表版本", "OptionsForm", "Label_Snd_Style", Lanuage_Now)
    .Label_SystemRoot.Caption = Load_Lanuage("操作系统所在位置", "OptionsForm", "Label_SystemRoot", Lanuage_Now)
    .Label_SysPath.Caption = Load_Lanuage("默认生成何种环境变量", "OptionsForm", "Label_SysPath", Lanuage_Now)
    .Frame_Soft_Glass.Caption = Load_Lanuage("本程序显示风格", "OptionsForm", "Frame_Soft_Glass", Lanuage_Now)
    .Aero_Normal.Caption = Load_Lanuage("普通", "OptionsForm", "Aero_Normal", Lanuage_Now)
    .Aero_Glass.Caption = Load_Lanuage("Aero全玻璃", "OptionsForm", "Aero_Glass", Lanuage_Now)
    .Frame_AutoPaper.Caption = Load_Lanuage("传送壁纸列表到“自动更换壁纸”程序", "OptionsForm", "Frame_AutoPaper", Lanuage_Now)
    .Option_AutoPaper_Y.Caption = Load_Lanuage("是", "OptionsForm", "Option_AutoPaper_Y", Lanuage_Now)
    .Option_AutoPaper_N.Caption = Load_Lanuage("否", "OptionsForm", "Option_AutoPaper_N", Lanuage_Now)
    .Option_AutoPaper_A.Caption = Load_Lanuage("询问", "OptionsForm", "Option_AutoPaper_A", Lanuage_Now)
    .Label_Aplha_Back_Color.Caption = Load_Lanuage("颜色预览背景颜色", "OptionsForm", "Label_Aplha_Back_Color", Lanuage_Now)
    .Check_frmLoad.Caption = Load_Lanuage("启动程序时不出现引导窗口", "OptionsForm", "Check_frmLoad", Lanuage_Now)
    .Command_Done.Caption = Load_Lanuage("OK", "OptionsForm", "Command_Done", Lanuage_Now)
    .Command_Cancel.Caption = Load_Lanuage("Cancel", "OptionsForm", "Command_Cancel", Lanuage_Now)
    .Command_Aply.Caption = Load_Lanuage("Aply", "OptionsForm", "Command_Aply", Lanuage_Now)
End With
'主窗口
With Main
    .Caption = Load_Lanuage("枫谷主题", "info", "AppName", Lanuage_Now) & " V" & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision
    .Check_ver.Caption = Load_Lanuage("检查更新", "Main", "Check_ver", Lanuage_Now)
    .Command_about.Caption = Load_Lanuage("关于", "Main", "Command_about", Lanuage_Now)
    .Command_Options.Caption = Load_Lanuage("设置", "Main", "Command_Options", Lanuage_Now)
    .Command_exit.Caption = Load_Lanuage("退出", "Main", "Command_exit", Lanuage_Now)
    .Option_Main_Tab(0).Caption = Load_Lanuage("选择主题文件", "Main", "Option_Main_Tab0", Lanuage_Now)
    .Option_Main_Tab(1).Caption = Load_Lanuage("手动应用", "Main", "Option_Main_Tab1", Lanuage_Now)
    .Option_Main_Tab(2).Caption = Load_Lanuage("编辑主题文件", "Main", "Option_Main_Tab2", Lanuage_Now)
    .Command_Guide.Caption = Load_Lanuage("导出向导", "Main", "Option_Main_Tab3", Lanuage_Now)
    '选择主题
    .Label_Help_Select_Theme.Caption = Load_Lanuage("下面列表内是您系统中已经安装的主题|请选择您需要应用的或者编辑的主题", "Main", "Help_Select_Theme", Lanuage_Now)
    .Command_Choose_Aply_Theme.Caption = Load_Lanuage("应用到系统", "Main", "Command_Choose_Aply_Theme", Lanuage_Now)
    .Command_Choose_Add_Theme.Caption = Load_Lanuage("添加列表中没有的主题", "Main", "Command_Choose_Add_Theme", Lanuage_Now)
    .Command_Choose_Edit_Theme.Caption = Load_Lanuage("编辑该主题", "Main", "Command_Choose_Edit_Theme", Lanuage_Now)
    .Command_Choose_Refresh_Theme.Caption = Load_Lanuage("刷新列表", "Main", "Command_Choose_Refresh_Theme", Lanuage_Now)
    .Command_Down_More_Theme.Caption = Load_Lanuage("获取更多主题", "Main", "Command_Down_More_Theme", Lanuage_Now)
    '手动应用
    .Label_mss_indro.Caption = Load_Lanuage("自动应用视觉风格文件可能会重启主题失败，可尝试多点几次。(需管理员权限启动本程序）|如果一直没有应用成功，请检查您是否破解了主题，或者您选择的视觉风格文件的操作系统是否对应", "Main", "Help_Aply_By_Hand", Lanuage_Now)
    .Command_ico_hand.Caption = Load_Lanuage("更改桌面图标", "Main", "Command_ico_hand", Lanuage_Now)
    .Command_cur_hand.Caption = Load_Lanuage("更改鼠标指针", "Main", "Command_cur_hand", Lanuage_Now)
    .Command_snd_hand.Caption = Load_Lanuage("更改系统音效", "Main", "Command_snd_hand", Lanuage_Now)
    .Command_paper_hand.Caption = Load_Lanuage("更改桌面壁纸", "Main", "Command_paper_hand", Lanuage_Now)
    .Command_window_hand.Caption = Load_Lanuage("更改窗体颜色和外观", "Main", "Command_window_hand", Lanuage_Now)
    .Command_glass_hand.Caption = Load_Lanuage("更改透明颜色", "Main", "Command_glass_hand", Lanuage_Now)
    .Command_individuation_hand.Caption = Load_Lanuage("打开个性化", "Main", "Command_individuation_hand", Lanuage_Now)
    .Command_scr_hand.Caption = Load_Lanuage("安装屏幕保护程序", "Main", "Command_scr_hand", Lanuage_Now)
    .Command_mss_hand.Caption = Load_Lanuage("修改视觉风格", "Main", "Command_mss_hand", Lanuage_Now)
    .Label_scr_hand.Caption = Load_Lanuage("屏幕保护程序文件", "Public", "CommonDialog_Scr_Filter", Lanuage_Now)
    .Label_mss_hand.Caption = Load_Lanuage("视觉风格文件", "Public", "CommonDialog_Mss_Filter", Lanuage_Now)
    .Command_scr_open.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
    .Command_mss_open.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
    '编辑
        '主体
        .Edit_Panel_Tab(0).Caption = Load_Lanuage("主题信息", "Main", "Edit_Panel_Theme_info_Caption", Lanuage_Now)
        .Edit_Panel_Tab(1).Caption = Load_Lanuage("视觉风格", "Main", "Edit_Panel_Mss_Caption", Lanuage_Now)
        .Edit_Panel_Tab(2).Caption = Load_Lanuage("桌面壁纸", "Main", "Edit_Panel_Paper_Caption", Lanuage_Now)
        .Edit_Panel_Tab(3).Caption = Load_Lanuage("桌面图标", "Main", "Edit_Panel_Icon_Caption", Lanuage_Now)
        .Edit_Panel_Tab(4).Caption = Load_Lanuage("鼠标指针", "Main", "Edit_Panel_Curson_Caption", Lanuage_Now)
        .Edit_Panel_Tab(5).Caption = Load_Lanuage("系统音效", "Main", "Edit_Panel_Sound_Caption", Lanuage_Now)
        .Edit_Panel_Tab(6).Caption = Load_Lanuage("屏幕保护程序", "Main", "Edit_Panel_Scr_Caption", Lanuage_Now)
        
        .Edit_Panel_Frame(0).Caption = Load_Lanuage("主题信息", "Main", "Edit_Panel_Theme_info_Caption", Lanuage_Now)
        .Edit_Panel_Frame(1).Caption = Load_Lanuage("视觉风格", "Main", "Edit_Panel_Mss_Caption", Lanuage_Now)
        .Edit_Panel_Frame(2).Caption = Load_Lanuage("桌面壁纸", "Main", "Edit_Panel_Paper_Caption", Lanuage_Now)
        .Edit_Panel_Frame(3).Caption = Load_Lanuage("桌面图标", "Main", "Edit_Panel_Icon_Caption", Lanuage_Now)
        .Edit_Panel_Frame(4).Caption = Load_Lanuage("鼠标指针", "Main", "Edit_Panel_Curson_Caption", Lanuage_Now)
        .Edit_Panel_Frame(5).Caption = Load_Lanuage("系统音效", "Main", "Edit_Panel_Sound_Caption", Lanuage_Now)
        .Edit_Panel_Frame(6).Caption = Load_Lanuage("屏幕保护程序", "Main", "Edit_Panel_Scr_Caption", Lanuage_Now)
        .Command_Aply_Now.Caption = Load_Lanuage("测试应用效果", "Main", "Command_Aply_Now", Lanuage_Now)
            '主题信息
            .Label_TnameC.Caption = Load_Lanuage("主题显示名称", "Main", "Label_Tname_Display", Lanuage_Now)
            .Label_TnameE.Caption = Load_Lanuage("主题文件名称", "Main", "Label_Tname_File", Lanuage_Now)
            .Label_maker.Caption = Load_Lanuage("主题制作者", "Main", "Label_maker", Lanuage_Now)
            .Label_maker_web.Caption = Load_Lanuage("网址或个人主页", "Main", "Label_maker_web", Lanuage_Now)
            .Label_Maker_Introduce.Caption = Load_Lanuage("其他版权信息或说明", "Main", "Label_Maker_Introduce", Lanuage_Now)
            .Label_Tlogo.Caption = Load_Lanuage("主题LOGO:", "Main", "Label_Tlogo", Lanuage_Now)
            .Label_Logo_Preview.Caption = Load_Lanuage("预览:", "Main", "Label_Logo_Preview", Lanuage_Now)
            .Label_logo_help.Caption = Load_Lanuage("推荐可以任意透明的PNG格式|LOGO最大显示为240×80像素（在256×256状态下）|因此请不要太大", "Main", "Label_logo_help", Lanuage_Now)
            .Command_Tlogo.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            '视觉风格
            .Frame_select_mss.Caption = Load_Lanuage("风格选择", "Main", "Frame_select_mss", Lanuage_Now)
            .mss_Aero.Caption = Load_Lanuage("Aero", "Main", "mss_Aero", Lanuage_Now)
            .Check_Alpha.Caption = Load_Lanuage("开启透明", "Main", "Check_Alpha", Lanuage_Now)
            .mss_Basic.Caption = Load_Lanuage("Basic", "Main", "mss_Basic", Lanuage_Now)
            .mss_Classic.Caption = Load_Lanuage("Classic", "Main", "mss_Classic", Lanuage_Now)
            .Command_getcolor.Caption = Load_Lanuage("一键取色", "Main", "Command_getcolor", Lanuage_Now)
            .System_Color_Tab(0).Caption = Load_Lanuage("可视化风格调节", "Main", "System_Color_Tab1", Lanuage_Now)
            .System_Color_Frame(0).Caption = .System_Color_Tab(0).Caption
            .System_Color_Tab(1).Caption = Load_Lanuage("窗口颜色和外观", "Main", "System_Color_Tab2", Lanuage_Now)
            .System_Color_Frame(1).Caption = .System_Color_Tab(1).Caption
            .Label_mss.Caption = Load_Lanuage("视觉风格文件", "Main", "Label_mss", Lanuage_Now)
            .Command_mss.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            .Label_ColorizationColor.Caption = Load_Lanuage("主颜色", "Main", "Label_ColorizationColor", Lanuage_Now)
            .Label_ColorizationColor_alpha.Caption = Load_Lanuage("主颜色透明度", "Main", "Label_ColorizationColor_alpha", Lanuage_Now)
            .Label_ColorizationColorBalance.Caption = Load_Lanuage("主颜色平衡", "Main", "Label_ColorizationColorBalance", Lanuage_Now)
            .Label_ColorizationAfterglow.Caption = Load_Lanuage("发光颜色", "Main", "Label_ColorizationAfterglow", Lanuage_Now)
            .Label_ColorizationAfterglow_alpha.Caption = Load_Lanuage("发光颜色透明度", "Main", "Label_ColorizationAfterglow_alpha", Lanuage_Now)
            .Label_ColorizationAfterglowBalance.Caption = Load_Lanuage("发光颜色平衡", "Main", "Label_ColorizationAfterglowBalance", Lanuage_Now)
            .Label_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("Aero条纹数量", "Main", "Label_ColorizationGlassReflectionIntensity", Lanuage_Now)
            .Label_ColorizationBlurBalance.Caption = Load_Lanuage("模糊平衡", "Main", "Label_ColorizationBlurBalance", Lanuage_Now)
            .Color_Warn.Caption = Load_Lanuage("自己编辑颜色可能导致一些奇怪的颜色，请使用一键取色工具", "Main", "Color_Warn", Lanuage_Now)
            .Label_Classic_Style.Caption = Load_Lanuage("经典风格预设: ", "Main", "Label_Classic_Style", Lanuage_Now)
            .Check_insert_system_color.Caption = Load_Lanuage("将自定义颜色加入到保存的主题或BAT文件中。（不选则为该风格系统默认值）·", "Main", "Check_insert_system_color", Lanuage_Now)
            '壁纸
            .Label_paper_index.Caption = Load_Lanuage("主壁纸文件:", "Main", "Label_paper_index", Lanuage_Now)
            .Command_paper.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            .Label_paper_style.Caption = Load_Lanuage("壁纸显示模式:", "Main", "Label_paper_style", Lanuage_Now)
            .Label_paper_change_time.Caption = Load_Lanuage("幻灯片切换时间:", "Main", "Label_paper_change_time", Lanuage_Now)
            .Check_paper_change.Caption = Load_Lanuage("无序切换", "Main", "Check_paper_change", Lanuage_Now)
            .Label_paper_files.Caption = Load_Lanuage("壁纸幻灯片文件夹:", "Main", "Label_paper_files", Lanuage_Now)
            .Command_paper_files.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            .Papers_Edit_Allow.Caption = Load_Lanuage("允许编辑图片列表", "Main", "Papers_Edit_Allow", Lanuage_Now)
            .Papers_Edit_Select_All.Caption = Load_Lanuage("选中全部", "Main", "Papers_Edit_Select_All", Lanuage_Now)
            .Papers_Edit_Clear.Caption = Load_Lanuage("全部不选", "Main", "Papers_Edit_Clear", Lanuage_Now)
            '图标
            For i = 0 To 5
                .Command_icon(i).Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            Next
            '鼠标指针
            .Label_cur.Caption = Load_Lanuage("方案名称", "Main", "Label_cur", Lanuage_Now)
            .Command_cur.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            .cur_default.Caption = Load_Lanuage("使用默认值(Win7)", "Main", "cur_default", Lanuage_Now)
            '音效
            .Check_snd.Caption = Load_Lanuage("使用已存在的方案", "Main", "Check_snd", Lanuage_Now)
            .Labe_sound_name_C.Caption = Load_Lanuage("声音方案中文", "Main", "Labe_sound_name_C", Lanuage_Now)
            .Label_sound_name_E.Caption = Load_Lanuage("声音方案英文简写", "Main", "Label_sound_name_E", Lanuage_Now)
            .sound_Play.Caption = Load_Lanuage("试听", "Main", "sound_Play", Lanuage_Now)
            .sound_Stop.Caption = Load_Lanuage("停止", "Main", "sound_Stop", Lanuage_Now)
            .Command_sound.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
            '屏幕保护程序
            .Label_scr_url.Caption = Load_Lanuage("屏保文件：（不启用屏保留空即可）", "Main", "Label_scr_url", Lanuage_Now)
            .Label_scr_wait.Caption = Load_Lanuage("等待:", "Main", "Label_scr_wait", Lanuage_Now)
            .Label_scr_wait_min.Caption = Load_Lanuage("分钟", "Main", "Label_scr_wait_min", Lanuage_Now)
            .Command_scr.Caption = Load_Lanuage("浏览", "Public", "Command_Select", Lanuage_Now)
End With
End Sub
