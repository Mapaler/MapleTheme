Attribute VB_Name = "读取文件"
Option Explicit

'读取theme文件
'读取ini如果想要用判断值，必须先判断不为空然后再判断，不然将会导致停止（或者是循环？）。
'读取的字符串后面跟有一些null，需要用一个文本框TEMP一下再存入变量
Public Sub Load_theme(ByVal Load_Url As String)
Dim i%, j%
Dim n%
Dim Theme_Ver%

    If GetFromIni("CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", "DefaultValue", Load_Url) <> "" Or GetFromIni("CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) <> "" Then
        If GetFromIni("VisualStyles", "ColorizationColor", Load_Url) <> "" Then
'            MsgBox "您选择的可能是Windows Vista主题文件。", 48, "提醒"
            Theme_Ver = 6
        Else
'            MsgBox "您选择的可能是Windows XP主题文件。", 48, "提醒"
        Theme_Ver = 5
        End If
    ElseIf GetFromIni("MasterThemeSelector", "MTSM", Load_Url) = "RJSPBS" Then
'        MsgBox "您选择的可能是Windows 8主题文件。", 48, "提醒"
        Theme_Ver = 8
    Else
        Theme_Ver = 7
    End If
    
    Dim Line() As String
    Open Load_Url For Input As #1
    Do Until EOF(1)
        ReDim Preserve Line(n)
        Line Input #1, Line(n)
        n = n + 1
'        If n >= 5 Then Exit Do '只读取前五行，然后退出循环
    Loop
    
        '樱茶生成器老版主题，读取进来只有一行
        If InStr(Line(0), "此主题文件由樱茶Win7主题生成器生成") <> 0 And UBound(Line) = 0 Then
'            MsgBox "欢迎您使用樱茶Win7主题生成器生成主题，但是您选择的是由樱茶Win7主题生成器2.02以前版本生成的主题文件。" & vbCrLf & "其换行编码非Windows方式，可能导致部分数据不能正确读取", 64, "提醒"
            If InStr(Line(0), ";作者：") <> 0 Then
                Main.Maker_Name.text = Mid$(Line(0), InStr(Line(0), ";作者：") + 4, InStr(InStr(Line(0), ";作者："), Line(0), vbLf) - (InStr(Line(0), ";作者：") + 4))
            End If
            If InStr(Line(0), ";网址：") <> 0 Then
                Main.Maker_Web_Url.text = Mid$(Line(0), InStr(Line(0), ";网址：") + 4, InStr(InStr(Line(0), ";网址："), Line(0), vbLf) - (InStr(Line(0), ";网址：") + 4))
            End If
        Else
    
            For i = 0 To UBound(Line)
                If InStr(Line(i), ";作者：") <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";作者：") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";作者:") <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";作者:") + 4)
                    Exit For
                ElseIf InStr(1, Line(i), ";by ", 1) <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";by ") + 4)
                    Exit For
                End If
            Next
            
            For i = 0 To UBound(Line)
                If InStr(Line(i), ";网址：") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";网址：") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";网址:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";网址:") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";个人主页：") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";个人主页：") + 6)
                    Exit For
                ElseIf InStr(Line(i), ";个人主页:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";个人主页:") + 6)
                    Exit For
                ElseIf InStr(Line(i), ";URL：") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";URL：") + 5)
                    Exit For
                ElseIf InStr(Line(i), ";URL:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";URL:") + 5)
                    Exit For
                End If
            Next
        End If
    
    Close #1

    '主题信息部分
    Main.T_name_E.text = Mid(Load_Url, InStrRev(Load_Url, "\") + 1, InStrRev(Load_Url, ".") - InStrRev(Load_Url, "\") - 1) '将文件名转换为英文名
    If GetFromIni("Theme", "DisplayName", Load_Url) <> "" Then
        Main.T_name_C.text = (GetFromIni("Theme", "DisplayName", Load_Url))
    Else
        Main.T_name_C.text = Main.T_name_E.text
    End If
    Main.url_Tlogo.text = GetFromIni("Theme", "BrandImage", Load_Url)
    
    '图标
    If Theme_Ver > 6 Then
        Main.url_icon(1).text = GetFromIni("CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon", "DefaultValue", Load_Url) '我的文档
        Main.url_icon(2).text = GetFromIni("CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon", "DefaultValue", Load_Url) '网上邻居
    Else 'XP、Vista
        Main.url_icon(1).text = GetFromIni("CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", "DefaultValue", Load_Url) '我的文档
        Main.url_icon(2).text = GetFromIni("CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) '网上邻居
    End If
    Main.url_icon(0).text = GetFromIni("CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) '我的电脑
    Main.url_icon(3).text = GetFromIni("CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Empty", Load_Url) '回收站空
    Main.url_icon(4).text = GetFromIni("CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Full", Load_Url) '回收站满
    Main.url_icon(5).text = GetFromIni("CLSID\{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) 'IE
    
    '鼠标
    Main.name_cur.text = Get_dll_text(GetFromIni("Control Panel\Cursors", "DefaultValue", Load_Url))
    For i = 0 To 14
        SysCur(i, 0) = StripTerminator(GetFromIni("Control Panel\Cursors", SysCur(i, 1), Load_Url))
    Next
    '壁纸
    Main.url_paper.text = GetFromIni("Control Panel\Desktop", "Wallpaper", Load_Url)
    '判断壁纸显示方式
    If GetFromIni("Control Panel\Desktop", "TileWallpaper", Load_Url) <> "" Then
        If GetFromIni("Control Panel\Desktop", "TileWallpaper", Load_Url) = 1 Then
                Main.ImageCombo_paper_style.ComboItems("平铺").Selected = True
        ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) <> "0" Then
            If GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 0 Then
                Main.ImageCombo_paper_style.ComboItems("居中").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 2 Then
                Main.ImageCombo_paper_style.ComboItems("拉伸").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 6 Then
                Main.ImageCombo_paper_style.ComboItems("适应").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 10 Then
                Main.ImageCombo_paper_style.ComboItems("填充").Selected = True
            Else
                Main.ImageCombo_paper_style.ComboItems("填充").Selected = True
            End If
        Else
            Main.ImageCombo_paper_style.ComboItems("填充").Selected = True
        End If
    
    Else
        Main.ImageCombo_paper_style.ComboItems("填充").Selected = True
    End If
    '选择幻灯片切换时间
    If GetFromIni("Slideshow", "Interval", Load_Url) <> "" Then
        Dim Time_Temp As Long
        Time_Temp = GetFromIni("Slideshow", "Interval", Load_Url) / 1000
        If Time_Temp < 15 Then
            Main.Combo_Paper_Change_Time.ListIndex = 0     '10秒
        ElseIf Time_Temp >= 15 And Time_Temp < 45 Then
            Main.Combo_Paper_Change_Time.ListIndex = 1
        ElseIf Time_Temp >= 45 And Time_Temp < 120 Then
            Main.Combo_Paper_Change_Time.ListIndex = 2
        ElseIf Time_Temp >= 120 And Time_Temp < 240 Then
            Main.Combo_Paper_Change_Time.ListIndex = 3
        ElseIf Time_Temp >= 240 And Time_Temp < 450 Then
            Main.Combo_Paper_Change_Time.ListIndex = 4
        ElseIf Time_Temp >= 450 And Time_Temp < 750 Then
            Main.Combo_Paper_Change_Time.ListIndex = 5
        ElseIf Time_Temp >= 750 And Time_Temp < 1050 Then
            Main.Combo_Paper_Change_Time.ListIndex = 6
        ElseIf Time_Temp >= 1050 And Time_Temp < 1500 Then
            Main.Combo_Paper_Change_Time.ListIndex = 7
        ElseIf Time_Temp >= 1500 And Time_Temp < 2700 Then
            Main.Combo_Paper_Change_Time.ListIndex = 8
        ElseIf Time_Temp >= 2700 And Time_Temp < 5400 Then
            Main.Combo_Paper_Change_Time.ListIndex = 9
        ElseIf Time_Temp >= 5400 And Time_Temp < 9000 Then
            Main.Combo_Paper_Change_Time.ListIndex = 10
        ElseIf Time_Temp >= 9000 And Time_Temp < 16200 Then
            Main.Combo_Paper_Change_Time.ListIndex = 11
        ElseIf Time_Temp >= 16200 And Time_Temp < 32400 Then
            Main.Combo_Paper_Change_Time.ListIndex = 12
        ElseIf Time_Temp >= 32400 And Time_Temp < 64800 Then
            Main.Combo_Paper_Change_Time.ListIndex = 13
        ElseIf Time_Temp >= 64800 Then
            Main.Combo_Paper_Change_Time.ListIndex = 14     '一天
        Else
            Main.Combo_Paper_Change_Time.ListIndex = 8     '30分钟
        End If
    Else
        Main.Combo_Paper_Change_Time.ListIndex = 8     '30分钟
    End If
    
    If GetFromIni("Slideshow", "Shuffle", Load_Url) <> "" Then
        If GetFromIni("Slideshow", "Shuffle", Load_Url) = 1 Then
            Main.Check_paper_change.value = 1
        Else
            Main.Check_paper_change.value = 0
        End If
    End If
    
    Main.url_paper_files.text = GetFromIni("Slideshow", "ImagesRootPath", Load_Url) '壁纸集
'    Dim rootPathPapers As Collection
'    Set rootPathPapers = New Collection
'    If Dir(Main.url_paper_files.text, vbDirectory + vbHidden + vbSystem) <> "" Then
'        Call GetFileName(Main.url_paper_files.text, "bmp,jpg,jpeg,gif,png", rootPathPapers) '获取该文件夹下图片文件列表
'    End If
'    Debug.Print rootPathPapers.count
    '如果有编辑图片列表的话
'    If GetFromIni("Slideshow1", "Item0Path", Load_Url) <> "" Then
'        Main.TreeView_paper.Nodes.Clear '清除以前的节点
'        Main.ImageList_wallpapers.ListImages.Clear '清除以前的图片
'        Erase PaperFileName '先清空数组
'        Main.Papers_Edit_Allow.Value = 1 '进入编辑状态
'        '将文件都添加到数组
'
'        Dim lngfiles As Integer '文件个数
'        Dim pi, pj As Long '循环次数
'        Dim lastnum As Integer
'        lngfiles = 0
'        lastnum = 0
'        For pj = 1 To 32767
'            For pi = 1 To 32767
'                If GetFromIni("Slideshow" & pj, "Item" & pi - 1 & "Path", Load_Url) <> "" Then
'                    lngfiles = lngfiles + 1
'                    ReDim Preserve PaperFileName(1 To lngfiles)
'                    PaperFileName(pi + lastnum) = url_to_N(GetFromIni("Slideshow" & pj, "Item" & pi - 1 & "Path", Load_Url))
'                Else
'                    Exit For
'                End If
'            Next pi
'            lastnum = lastnum + pi - 1 '记录上之前的壁纸数
'            If GetFromIni("Slideshow" & pj, "Item0Path", Load_Url) = "" Then
'                Exit For
'            End If
'        Next pj
'        '将所有图片载入imagelist
'        For i = 1 To UBound(PaperFileName)
'            Main.Picture_paper_TEMP.Cls
'            Call PaintPng2(PaperFileName(i), Main.Picture_paper_TEMP.hdc, pWidth, pHeight)
'            Main.ImageList_wallpapers.ListImages.Add , , Main.Picture_paper_TEMP.image
'        Next
'        '添加listview节点
'        Dim Root As Node
'        With Main.TreeView_paper.Nodes
'            For i = 1 To UBound(PaperFileName)
'                Set Root = .Add(, , "p" & i, Mid$(PaperFileName(i), InStrRev(PaperFileName(i), "\") + 1), i)
'            Next
'        End With
'        For i = 1 To Main.TreeView_paper.Nodes.Count
'            Main.TreeView_paper.Nodes(i).Checked = True '壁纸全选中
'        Next
'    End If
    
    
    '声音
    Dim snd_not_empty As Boolean
    snd_not_empty = False
    n = 0
    For i = 0 To UBound(F_Sound)
'        If F_Sound(i) <> "" Then '因为从语言文件读取的会有一个空的不过目前我直接把读取的地方的空的消掉了
            For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
                Sound(2, n) = GetFromIni("AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n), "DefaultValue", Load_Url)
                If Sound(2, n) <> "" Then
                    Main.TreeView_Sound.Nodes("s" & n).Image = 2 '将有值的节点小喇叭变白
                    snd_not_empty = True
                Else
                    Main.TreeView_Sound.Nodes("s" & n).Image = 0 '将无值的节点小喇叭删掉
                End If
                n = n + 1
            Next j
'        End If
    Next i
    
    If snd_not_empty = True Then
        Main.Check_snd.value = 0
        If GetFromIni("Sounds", "SchemeName", Load_Url) <> "" Then
            Main.sound_name_C.text = Get_dll_text(GetFromIni("Sounds", "SchemeName", Load_Url))
        Else
            Main.sound_name_C.text = Main.T_name_C.text
        End If
        Main.sound_name_E.text = Main.T_name_E.text
        Main.Check_snd.value = 0
    Else
        Main.Check_snd.value = 1
        For i = 0 To UBound(Sound_Name)
            If Get_dll_text(GetFromIni("Sounds", "SchemeName", Load_Url)) = Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Names\" & Sound_Name(i), vbNullString)) Then
                Main.Combo_Sys_Snd.ListIndex = i
                Exit For
            Else
                Main.Combo_Sys_Snd.ListIndex = 0
            End If
        Next
    End If
    
    
    
    
    
    '屏幕保护
    Main.url_scr.text = GetFromIni("boot", "SCRNSAVE.EXE", Load_Url)
    '视觉风格
    Main.url_mss.text = url_to_P(GetFromIni("VisualStyles", "Path", Load_Url))
    Main.Value_ColorizationColor.text = GetFromIni("VisualStyles", "ColorizationColor", Load_Url)
    Main.Value_ColorizationAfterglow.text = Main.Value_ColorizationColor.text
    Main.Value_ColorizationColorBalance.text = 8
    Main.Value_ColorizationAfterglowBalance.text = 43
    Main.Value_ColorizationBlurBalance.text = 49
    Main.Value_ColorizationGlassReflectionIntensity.text = 50
    

    '判断视觉风格
    If GetFromIni("VisualStyles", "ColorStyle", Load_Url) <> "" Then
        Dim ColorStyletemp1 As String
        ColorStyletemp1 = Get_dll_text(GetFromIni("VisualStyles", "ColorStyle", Load_Url))
        If ColorStyletemp1 = Get_dll_text("@themeui.dll,-2027") Then
            '如果ColorStyle为NormalColor则判断是Basic还是Aero
            If GetFromIni("VisualStyles", "Composition", Load_Url) <> "" Then
                If GetFromIni("VisualStyles", "Composition", Load_Url) = 0 Then
'                    MsgBox "您当前选择的主题是Basic模式", 64, "提醒"
                    Main.mss_Basic.value = True 'Basic
                Else
                    Main.mss_Aero.value = True
                End If
            Else
                Main.mss_Aero.value = True
            End If
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-854") Then
'            MsgBox "您当前选择的主题是Windows经典", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S2").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-850") Then
'            MsgBox "您当前选择的主题是高对比度 #1", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S3").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-851") Then
'            MsgBox "您当前选择的主题是高对比度 #2", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S4").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-852") Then
'            MsgBox "您当前选择的主题是高对比黑色", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S5").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-853") Then
'            MsgBox "您当前选择的主题是高对比白色", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S6").Selected = True
        ElseIf GetFromIni("VisualStyles", "Size", Load_Url) = Get_dll_text("@themeui.dll,-853") Then
'            MsgBox "您当前选择的主题是自定义", 64, "提醒"
            Main.mss_Classic.value = True 'Windows经典
            Main.ImageCombo_Classic_Style.ComboItems("C_S1").Selected = True
        Else
'            MsgBox "ColorStyle识别失败，默认切换到Aero", 64, "提醒"
            Main.mss_Aero.value = True
        End If
    Else
        Main.mss_Aero.value = True
    End If
    '检查透明
    If GetFromIni("VisualStyles", "Transparency", Load_Url) <> "" Then
        If GetFromIni("VisualStyles", "Transparency", Load_Url) = 0 Then
            Main.Check_Alpha.value = 0 '不透明
        Else
            Main.Check_Alpha.value = 1 '透明
        End If
    Else
        Main.Check_Alpha.value = 1 '透明
    End If
    '系统颜色
    If GetFromIni("Control Panel\Colors", SysColors(0, 0), Load_Url) <> "" Then
        Main.Check_insert_system_color.value = 1
        For i = 0 To 30
            SysColors(i, 1) = GetFromIni("Control Panel\Colors", SysColors(i, 0), Load_Url)
        Next
        Main.ImageCombo_Classic_Style.ComboItems(1).Selected = True
    End If
End Sub
