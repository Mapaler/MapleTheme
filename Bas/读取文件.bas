Attribute VB_Name = "��ȡ�ļ�"
Option Explicit

'��ȡtheme�ļ�
'��ȡini�����Ҫ���ж�ֵ���������жϲ�Ϊ��Ȼ�����жϣ���Ȼ���ᵼ��ֹͣ��������ѭ��������
'��ȡ���ַ����������һЩnull����Ҫ��һ���ı���TEMPһ���ٴ������
Public Sub Load_theme(ByVal Load_Url As String)
Dim i%, j%
Dim n%
Dim Theme_Ver%

    If GetFromIni("CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", "DefaultValue", Load_Url) <> "" Or GetFromIni("CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) <> "" Then
        If GetFromIni("VisualStyles", "ColorizationColor", Load_Url) <> "" Then
'            MsgBox "��ѡ��Ŀ�����Windows Vista�����ļ���", 48, "����"
            Theme_Ver = 6
        Else
'            MsgBox "��ѡ��Ŀ�����Windows XP�����ļ���", 48, "����"
        Theme_Ver = 5
        End If
    ElseIf GetFromIni("MasterThemeSelector", "MTSM", Load_Url) = "RJSPBS" Then
'        MsgBox "��ѡ��Ŀ�����Windows 8�����ļ���", 48, "����"
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
'        If n >= 5 Then Exit Do 'ֻ��ȡǰ���У�Ȼ���˳�ѭ��
    Loop
    
        'ӣ���������ϰ����⣬��ȡ����ֻ��һ��
        If InStr(Line(0), "�������ļ���ӣ��Win7��������������") <> 0 And UBound(Line) = 0 Then
'            MsgBox "��ӭ��ʹ��ӣ��Win7�����������������⣬������ѡ�������ӣ��Win7����������2.02��ǰ�汾���ɵ������ļ���" & vbCrLf & "�任�б����Windows��ʽ�����ܵ��²������ݲ�����ȷ��ȡ", 64, "����"
            If InStr(Line(0), ";���ߣ�") <> 0 Then
                Main.Maker_Name.text = Mid$(Line(0), InStr(Line(0), ";���ߣ�") + 4, InStr(InStr(Line(0), ";���ߣ�"), Line(0), vbLf) - (InStr(Line(0), ";���ߣ�") + 4))
            End If
            If InStr(Line(0), ";��ַ��") <> 0 Then
                Main.Maker_Web_Url.text = Mid$(Line(0), InStr(Line(0), ";��ַ��") + 4, InStr(InStr(Line(0), ";��ַ��"), Line(0), vbLf) - (InStr(Line(0), ";��ַ��") + 4))
            End If
        Else
    
            For i = 0 To UBound(Line)
                If InStr(Line(i), ";���ߣ�") <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";���ߣ�") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";����:") <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";����:") + 4)
                    Exit For
                ElseIf InStr(1, Line(i), ";by ", 1) <> 0 Then
                    Main.Maker_Name.text = Mid$(Line(i), InStr(Line(i), ";by ") + 4)
                    Exit For
                End If
            Next
            
            For i = 0 To UBound(Line)
                If InStr(Line(i), ";��ַ��") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";��ַ��") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";��ַ:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";��ַ:") + 4)
                    Exit For
                ElseIf InStr(Line(i), ";������ҳ��") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";������ҳ��") + 6)
                    Exit For
                ElseIf InStr(Line(i), ";������ҳ:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";������ҳ:") + 6)
                    Exit For
                ElseIf InStr(Line(i), ";URL��") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";URL��") + 5)
                    Exit For
                ElseIf InStr(Line(i), ";URL:") <> 0 Then
                    Main.Maker_Web_Url.text = Mid$(Line(i), InStr(Line(i), ";URL:") + 5)
                    Exit For
                End If
            Next
        End If
    
    Close #1

    '������Ϣ����
    Main.T_name_E.text = Mid(Load_Url, InStrRev(Load_Url, "\") + 1, InStrRev(Load_Url, ".") - InStrRev(Load_Url, "\") - 1) '���ļ���ת��ΪӢ����
    If GetFromIni("Theme", "DisplayName", Load_Url) <> "" Then
        Main.T_name_C.text = (GetFromIni("Theme", "DisplayName", Load_Url))
    Else
        Main.T_name_C.text = Main.T_name_E.text
    End If
    Main.url_Tlogo.text = GetFromIni("Theme", "BrandImage", Load_Url)
    
    'ͼ��
    If Theme_Ver > 6 Then
        Main.url_icon(1).text = GetFromIni("CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon", "DefaultValue", Load_Url) '�ҵ��ĵ�
        Main.url_icon(2).text = GetFromIni("CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon", "DefaultValue", Load_Url) '�����ھ�
    Else 'XP��Vista
        Main.url_icon(1).text = GetFromIni("CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\DefaultIcon", "DefaultValue", Load_Url) '�ҵ��ĵ�
        Main.url_icon(2).text = GetFromIni("CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) '�����ھ�
    End If
    Main.url_icon(0).text = GetFromIni("CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) '�ҵĵ���
    Main.url_icon(3).text = GetFromIni("CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Empty", Load_Url) '����վ��
    Main.url_icon(4).text = GetFromIni("CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon", "Full", Load_Url) '����վ��
    Main.url_icon(5).text = GetFromIni("CLSID\{871C5380-42A0-1069-A2EA-08002B30309D}\DefaultIcon", "DefaultValue", Load_Url) 'IE
    
    '���
    Main.name_cur.text = Get_dll_text(GetFromIni("Control Panel\Cursors", "DefaultValue", Load_Url))
    For i = 0 To 14
        SysCur(i, 0) = StripTerminator(GetFromIni("Control Panel\Cursors", SysCur(i, 1), Load_Url))
    Next
    '��ֽ
    Main.url_paper.text = GetFromIni("Control Panel\Desktop", "Wallpaper", Load_Url)
    '�жϱ�ֽ��ʾ��ʽ
    If GetFromIni("Control Panel\Desktop", "TileWallpaper", Load_Url) <> "" Then
        If GetFromIni("Control Panel\Desktop", "TileWallpaper", Load_Url) = 1 Then
                Main.ImageCombo_paper_style.ComboItems("ƽ��").Selected = True
        ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) <> "0" Then
            If GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 0 Then
                Main.ImageCombo_paper_style.ComboItems("����").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 2 Then
                Main.ImageCombo_paper_style.ComboItems("����").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 6 Then
                Main.ImageCombo_paper_style.ComboItems("��Ӧ").Selected = True
            ElseIf GetFromIni("Control Panel\Desktop", "WallpaperStyle", Load_Url) = 10 Then
                Main.ImageCombo_paper_style.ComboItems("���").Selected = True
            Else
                Main.ImageCombo_paper_style.ComboItems("���").Selected = True
            End If
        Else
            Main.ImageCombo_paper_style.ComboItems("���").Selected = True
        End If
    
    Else
        Main.ImageCombo_paper_style.ComboItems("���").Selected = True
    End If
    'ѡ��õ�Ƭ�л�ʱ��
    If GetFromIni("Slideshow", "Interval", Load_Url) <> "" Then
        Dim Time_Temp As Long
        Time_Temp = GetFromIni("Slideshow", "Interval", Load_Url) / 1000
        If Time_Temp < 15 Then
            Main.Combo_Paper_Change_Time.ListIndex = 0     '10��
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
            Main.Combo_Paper_Change_Time.ListIndex = 14     'һ��
        Else
            Main.Combo_Paper_Change_Time.ListIndex = 8     '30����
        End If
    Else
        Main.Combo_Paper_Change_Time.ListIndex = 8     '30����
    End If
    
    If GetFromIni("Slideshow", "Shuffle", Load_Url) <> "" Then
        If GetFromIni("Slideshow", "Shuffle", Load_Url) = 1 Then
            Main.Check_paper_change.value = 1
        Else
            Main.Check_paper_change.value = 0
        End If
    End If
    
    Main.url_paper_files.text = GetFromIni("Slideshow", "ImagesRootPath", Load_Url) '��ֽ��
'    Dim rootPathPapers As Collection
'    Set rootPathPapers = New Collection
'    If Dir(Main.url_paper_files.text, vbDirectory + vbHidden + vbSystem) <> "" Then
'        Call GetFileName(Main.url_paper_files.text, "bmp,jpg,jpeg,gif,png", rootPathPapers) '��ȡ���ļ�����ͼƬ�ļ��б�
'    End If
'    Debug.Print rootPathPapers.count
    '����б༭ͼƬ�б�Ļ�
'    If GetFromIni("Slideshow1", "Item0Path", Load_Url) <> "" Then
'        Main.TreeView_paper.Nodes.Clear '�����ǰ�Ľڵ�
'        Main.ImageList_wallpapers.ListImages.Clear '�����ǰ��ͼƬ
'        Erase PaperFileName '���������
'        Main.Papers_Edit_Allow.Value = 1 '����༭״̬
'        '���ļ�����ӵ�����
'
'        Dim lngfiles As Integer '�ļ�����
'        Dim pi, pj As Long 'ѭ������
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
'            lastnum = lastnum + pi - 1 '��¼��֮ǰ�ı�ֽ��
'            If GetFromIni("Slideshow" & pj, "Item0Path", Load_Url) = "" Then
'                Exit For
'            End If
'        Next pj
'        '������ͼƬ����imagelist
'        For i = 1 To UBound(PaperFileName)
'            Main.Picture_paper_TEMP.Cls
'            Call PaintPng2(PaperFileName(i), Main.Picture_paper_TEMP.hdc, pWidth, pHeight)
'            Main.ImageList_wallpapers.ListImages.Add , , Main.Picture_paper_TEMP.image
'        Next
'        '���listview�ڵ�
'        Dim Root As Node
'        With Main.TreeView_paper.Nodes
'            For i = 1 To UBound(PaperFileName)
'                Set Root = .Add(, , "p" & i, Mid$(PaperFileName(i), InStrRev(PaperFileName(i), "\") + 1), i)
'            Next
'        End With
'        For i = 1 To Main.TreeView_paper.Nodes.Count
'            Main.TreeView_paper.Nodes(i).Checked = True '��ֽȫѡ��
'        Next
'    End If
    
    
    '����
    Dim snd_not_empty As Boolean
    snd_not_empty = False
    n = 0
    For i = 0 To UBound(F_Sound)
'        If F_Sound(i) <> "" Then '��Ϊ�������ļ���ȡ�Ļ���һ���յĲ���Ŀǰ��ֱ�ӰѶ�ȡ�ĵط��Ŀյ�������
            For j = 0 To Main.TreeView_Sound.Nodes("F_" & F_Sound(i)).Children - 1
                Sound(2, n) = GetFromIni("AppEvents\Schemes\Apps\" & F_Sound(i) & "\" & Sound(0, n), "DefaultValue", Load_Url)
                If Sound(2, n) <> "" Then
                    Main.TreeView_Sound.Nodes("s" & n).Image = 2 '����ֵ�Ľڵ�С���ȱ��
                    snd_not_empty = True
                Else
                    Main.TreeView_Sound.Nodes("s" & n).Image = 0 '����ֵ�Ľڵ�С����ɾ��
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
    
    
    
    
    
    '��Ļ����
    Main.url_scr.text = GetFromIni("boot", "SCRNSAVE.EXE", Load_Url)
    '�Ӿ����
    Main.url_mss.text = url_to_P(GetFromIni("VisualStyles", "Path", Load_Url))
    Main.Value_ColorizationColor.text = GetFromIni("VisualStyles", "ColorizationColor", Load_Url)
    Main.Value_ColorizationAfterglow.text = Main.Value_ColorizationColor.text
    Main.Value_ColorizationColorBalance.text = 8
    Main.Value_ColorizationAfterglowBalance.text = 43
    Main.Value_ColorizationBlurBalance.text = 49
    Main.Value_ColorizationGlassReflectionIntensity.text = 50
    

    '�ж��Ӿ����
    If GetFromIni("VisualStyles", "ColorStyle", Load_Url) <> "" Then
        Dim ColorStyletemp1 As String
        ColorStyletemp1 = Get_dll_text(GetFromIni("VisualStyles", "ColorStyle", Load_Url))
        If ColorStyletemp1 = Get_dll_text("@themeui.dll,-2027") Then
            '���ColorStyleΪNormalColor���ж���Basic����Aero
            If GetFromIni("VisualStyles", "Composition", Load_Url) <> "" Then
                If GetFromIni("VisualStyles", "Composition", Load_Url) = 0 Then
'                    MsgBox "����ǰѡ���������Basicģʽ", 64, "����"
                    Main.mss_Basic.value = True 'Basic
                Else
                    Main.mss_Aero.value = True
                End If
            Else
                Main.mss_Aero.value = True
            End If
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-854") Then
'            MsgBox "����ǰѡ���������Windows����", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S2").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-850") Then
'            MsgBox "����ǰѡ��������Ǹ߶Աȶ� #1", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S3").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-851") Then
'            MsgBox "����ǰѡ��������Ǹ߶Աȶ� #2", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S4").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-852") Then
'            MsgBox "����ǰѡ��������Ǹ߶ԱȺ�ɫ", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S5").Selected = True
        ElseIf ColorStyletemp1 = Get_dll_text("@themeui.dll,-853") Then
'            MsgBox "����ǰѡ��������Ǹ߶ԱȰ�ɫ", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S6").Selected = True
        ElseIf GetFromIni("VisualStyles", "Size", Load_Url) = Get_dll_text("@themeui.dll,-853") Then
'            MsgBox "����ǰѡ����������Զ���", 64, "����"
            Main.mss_Classic.value = True 'Windows����
            Main.ImageCombo_Classic_Style.ComboItems("C_S1").Selected = True
        Else
'            MsgBox "ColorStyleʶ��ʧ�ܣ�Ĭ���л���Aero", 64, "����"
            Main.mss_Aero.value = True
        End If
    Else
        Main.mss_Aero.value = True
    End If
    '���͸��
    If GetFromIni("VisualStyles", "Transparency", Load_Url) <> "" Then
        If GetFromIni("VisualStyles", "Transparency", Load_Url) = 0 Then
            Main.Check_Alpha.value = 0 '��͸��
        Else
            Main.Check_Alpha.value = 1 '͸��
        End If
    Else
        Main.Check_Alpha.value = 1 '͸��
    End If
    'ϵͳ��ɫ
    If GetFromIni("Control Panel\Colors", SysColors(0, 0), Load_Url) <> "" Then
        Main.Check_insert_system_color.value = 1
        For i = 0 To 30
            SysColors(i, 1) = GetFromIni("Control Panel\Colors", SysColors(i, 0), Load_Url)
        Next
        Main.ImageCombo_Classic_Style.ComboItems(1).Selected = True
    End If
End Sub
