Attribute VB_Name = "�ػ���������"
Option Explicit
Public Function Load_Lanuage(ByVal Now_Show As String, ByVal strSectionHeader As String, ByVal strVariableName As String, Optional ByVal Change_Lanuage_Now As Integer = -1) As String
Dim text_temp() As String, text_temp2 As String
Dim i As Integer

If Change_Lanuage_Now = -1 Then
    Change_Lanuage_Now = Lanuage_Now
End If

    Erase text_temp '���������
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
'������������
With frmLoad
    .Caption = Load_Lanuage("������� - ѡ����������", "Load", "Caption", Lanuage_Now)
    .Frame_Basic.Caption = Load_Lanuage("Window7��ͥ��ͨ��Ӧ������", "Load", "Frame_Basic", Lanuage_Now)
    .Frame_Edit.Caption = Load_Lanuage("�༭ / ����Windows����", "Load", "Frame_Edit", Lanuage_Now)
    .Command_Open_Control.Caption = Load_Lanuage("�ֶ�Ӧ������", "Load", "Command_Open_Control", Lanuage_Now)
    .Command_theme_to_Bat.Caption = Load_Lanuage("�Զ�Ӧ�����⵽ϵͳ", "Load", "Command_theme_to_Bat", Lanuage_Now)
    .Command_Edit.Caption = Load_Lanuage("�򿪱༭��", "Load", "Command_Edit", Lanuage_Now)
    .Check_frmLoad.Caption = Load_Lanuage("�´β��ٳ��ֱ�����", "Load", "Check_frmLoad", Lanuage_Now)
End With
'����
With frmAbout
    .Caption = Load_Lanuage("����", "About", "Caption", Lanuage_Now) & " " & Load_Lanuage("�������", "info", "AppName", Lanuage_Now)
    .lblVersion.Caption = Load_Lanuage("�汾", "About", "Version", Lanuage_Now) & " " & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision
    .lblTitle.Caption = Load_Lanuage("�������", "info", "AppName", Lanuage_Now)
    .cmdOk.Caption = Load_Lanuage("ȷ��", "About", "cmdOK", Lanuage_Now)
    .cmdVisitVeb.Caption = Load_Lanuage("���ʹ���", "About", "cmdVisitVeb", Lanuage_Now)
    .lblDescription.Caption = Load_Lanuage("���������Զ��ռ���Դ�������ļ����������⣬����Ϊ�Ѿ���װ���������ɼ�ͥ�氲װBAT|�����������ڣ�|����������������ʹ��������Win7��ͥ��ͨ��ʹ�������BAT|ʹ�����������н�û�����Win7��ͥ�氲װBAT���������ɰ�װBAT", "About", "Description", Lanuage_Now)
    .lblDisclaimer.Caption = Load_Lanuage("����������Ȩ��ʹ��Ȩ���Ƚ�������", "About", "Disclaimer", Lanuage_Now)
End With
'һ��ȡɫ����
With Get_color
    .Caption = Load_Lanuage("һ��ȡɫ����", "Get_color", "Caption", Lanuage_Now)
    .freshen.Caption = Load_Lanuage("ˢ��", "Get_color", "freshen", Lanuage_Now)
    .freshen2.Caption = Load_Lanuage("ˢ��", "Get_color", "freshen", Lanuage_Now)
    .add_all.Caption = Load_Lanuage("����ȫ��", "Get_color", "add_all", Lanuage_Now)
    .add_all2.Caption = Load_Lanuage("����ȫ��", "Get_color", "add_all", Lanuage_Now)
    .Command_mss.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationColor.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationColorBalance.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationAfterglow.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationAfterglowBalance.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_ColorizationBlurBalance.Caption = Load_Lanuage("�����", "Get_color", "add_one", Lanuage_Now)
    .Command_glass.Caption = Load_Lanuage("��͸����ɫ���", "Get_color", "Command_glass", Lanuage_Now)
    .Command_window.Caption = Load_Lanuage("�򿪴�����ɫ��������", "Get_color", "Command_window", Lanuage_Now)
    
    .Label_mss.Caption = Load_Lanuage("�Ӿ�����ļ�", "Main", "Label_mss", Lanuage_Now)
    .Label_ColorizationColor.Caption = Load_Lanuage("����ɫ", "Main", "Label_ColorizationColor", Lanuage_Now)
    .Label_ColorizationColorBalance.Caption = Load_Lanuage("����ɫƽ��", "Main", "Label_ColorizationColorBalance", Lanuage_Now)
    .Label_ColorizationAfterglow.Caption = Load_Lanuage("������ɫ", "Main", "Label_ColorizationAfterglow", Lanuage_Now)
    .Label_ColorizationAfterglowBalance.Caption = Load_Lanuage("������ɫƽ��", "Main", "Label_ColorizationAfterglowBalance", Lanuage_Now)
    .Label_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("Aero��������", "Main", "Label_ColorizationGlassReflectionIntensity", Lanuage_Now)
    .Label_ColorizationBlurBalance.Caption = Load_Lanuage("ģ��ƽ��", "Main", "Label_ColorizationBlurBalance", Lanuage_Now)

    .Label_help.Caption = Load_Lanuage("ʹ�÷�����|��ȷ������Aero����£�Win7 HomeBasic�½���Windows 7 Standard����ʹ��Windows�Դ��ĸ��Ի�����ħ����AeroЧ�����ڣ����ں������������ɫ��������|Ȼ���������ߵġ�ˢ�¡����ͻ�ȥ������ǰ������ֵ��Ȼ�������Ҫѡ����뵽�����򴰿�����ȥ��|�ұ�����ϵͳ������ɫ���ã��ǵ���Classic����ʹ�á�", "Get_color", "help", Lanuage_Now)
End With
'������
With CreatGuide
    .Caption = Load_Lanuage("����������", "CreatGuide", "Caption", Lanuage_Now)
    .cmdLast.Caption = Load_Lanuage("��һ��", "CreatGuide", "cmdLast", Lanuage_Now)
    .cmdNext.Caption = Load_Lanuage("��һ��", "CreatGuide", "cmdNext", Lanuage_Now)
    .cmdOk.Caption = Load_Lanuage("���", "CreatGuide", "cmdOk", Lanuage_Now)
    .cmdCancel.Caption = Load_Lanuage("ȡ��", "CreatGuide", "cmdCancel", Lanuage_Now)
    
    .Frame_Files.Caption = Load_Lanuage("��ѡ�����ɺ����ļ�", "CreatGuide", "Files", Lanuage_Now)
    .Option_File(0).Caption = Load_Lanuage("Windows Theme�ļ�", "CreatGuide", "Option_File_Theme", Lanuage_Now)
    .Option_File(1).Caption = Load_Lanuage("Bat�ļ�������Win7��ͥ��ͨ�棩", "CreatGuide", "Option_File_Bat", Lanuage_Now)
    
    .Frame_Theme_Ver.Caption = Load_Lanuage("��ѡ������Ҫ���ɵİ汾", "CreatGuide", "Frame_Theme_Ver", Lanuage_Now)
    .Option_Theme_Ver(0).Caption = Load_Lanuage("Windowsͨ�ã�ע���Ӿ�����ļ�����ͨ�ã�", "CreatGuide", "Option_Theme_Ver_All", Lanuage_Now)
    .Option_Theme_Ver(1).Caption = Load_Lanuage("Windows XP / 2003", "CreatGuide", "Option_Theme_Ver_XP", Lanuage_Now)
    .Option_Theme_Ver(2).Caption = Load_Lanuage("Windows Vista / 2008", "CreatGuide", "Option_Theme_Ver_Vista", Lanuage_Now)
    .Option_Theme_Ver(3).Caption = Load_Lanuage("Windows 7 / 2008 R2", "CreatGuide", "Option_Theme_Ver_7", Lanuage_Now)
    .Option_Theme_Ver(4).Caption = Load_Lanuage("Windows 8", "CreatGuide", "Option_Theme_Ver_8", Lanuage_Now)
    
    .Frame_BT_Color.Caption = Load_Lanuage("��ѡ�����ɵ�BAT�ļ��������뱳��ɫ", "CreatGuide", "Frame_BT_Color", Lanuage_Now)
    .Frame_BT_Color_Fore.Caption = Load_Lanuage("ǰ��ɫ", "CreatGuide", "Frame_BT_Color_Fore", Lanuage_Now)
    .Frame_BT_Color_Back.Caption = Load_Lanuage("����ɫ", "CreatGuide", "Frame_BT_Color_Back", Lanuage_Now)
End With
'����
With Options
    .Caption = Load_Lanuage("ѡ������", "OptionsForm", "Caption", Lanuage_Now)
    .Label_Lanuage.Caption = Load_Lanuage("�������/Lanuages", "OptionsForm", "Label_Lanuage", Lanuage_Now)
    .Command_Find_Lanuages.Caption = Load_Lanuage("��ȡ����Find More", "OptionsForm", "Command_Find_Lanuages", Lanuage_Now)
    .Frame_SystemTextShow.Caption = Load_Lanuage("����ϵͳ����", "OptionsForm", "Frame_SystemTextShow", Lanuage_Now)
    .SystemTextShow_Sys.Caption = Load_Lanuage("��ϵͳ��ȡ", "OptionsForm", "SystemTextShow_Sys", Lanuage_Now)
    .SystemTextShow_ini.Caption = Load_Lanuage("�������ļ���ȡ", "OptionsForm", "SystemTextShow_ini", Lanuage_Now)
    .Label_Snd_Style.Caption = Load_Lanuage("��Ч�б�汾", "OptionsForm", "Label_Snd_Style", Lanuage_Now)
    .Label_SystemRoot.Caption = Load_Lanuage("����ϵͳ����λ��", "OptionsForm", "Label_SystemRoot", Lanuage_Now)
    .Label_SysPath.Caption = Load_Lanuage("Ĭ�����ɺ��ֻ�������", "OptionsForm", "Label_SysPath", Lanuage_Now)
    .Frame_Soft_Glass.Caption = Load_Lanuage("��������ʾ���", "OptionsForm", "Frame_Soft_Glass", Lanuage_Now)
    .Aero_Normal.Caption = Load_Lanuage("��ͨ", "OptionsForm", "Aero_Normal", Lanuage_Now)
    .Aero_Glass.Caption = Load_Lanuage("Aeroȫ����", "OptionsForm", "Aero_Glass", Lanuage_Now)
    .Frame_AutoPaper.Caption = Load_Lanuage("���ͱ�ֽ�б����Զ�������ֽ������", "OptionsForm", "Frame_AutoPaper", Lanuage_Now)
    .Option_AutoPaper_Y.Caption = Load_Lanuage("��", "OptionsForm", "Option_AutoPaper_Y", Lanuage_Now)
    .Option_AutoPaper_N.Caption = Load_Lanuage("��", "OptionsForm", "Option_AutoPaper_N", Lanuage_Now)
    .Option_AutoPaper_A.Caption = Load_Lanuage("ѯ��", "OptionsForm", "Option_AutoPaper_A", Lanuage_Now)
    .Label_Aplha_Back_Color.Caption = Load_Lanuage("��ɫԤ��������ɫ", "OptionsForm", "Label_Aplha_Back_Color", Lanuage_Now)
    .Check_frmLoad.Caption = Load_Lanuage("��������ʱ��������������", "OptionsForm", "Check_frmLoad", Lanuage_Now)
    .Command_Done.Caption = Load_Lanuage("OK", "OptionsForm", "Command_Done", Lanuage_Now)
    .Command_Cancel.Caption = Load_Lanuage("Cancel", "OptionsForm", "Command_Cancel", Lanuage_Now)
    .Command_Aply.Caption = Load_Lanuage("Aply", "OptionsForm", "Command_Aply", Lanuage_Now)
End With
'������
With Main
    .Caption = Load_Lanuage("�������", "info", "AppName", Lanuage_Now) & " V" & App.Major & "." & App.Minor & App_Beta & " Build " & App.Revision
    .Check_ver.Caption = Load_Lanuage("������", "Main", "Check_ver", Lanuage_Now)
    .Command_about.Caption = Load_Lanuage("����", "Main", "Command_about", Lanuage_Now)
    .Command_Options.Caption = Load_Lanuage("����", "Main", "Command_Options", Lanuage_Now)
    .Command_exit.Caption = Load_Lanuage("�˳�", "Main", "Command_exit", Lanuage_Now)
    .Option_Main_Tab(0).Caption = Load_Lanuage("ѡ�������ļ�", "Main", "Option_Main_Tab0", Lanuage_Now)
    .Option_Main_Tab(1).Caption = Load_Lanuage("�ֶ�Ӧ��", "Main", "Option_Main_Tab1", Lanuage_Now)
    .Option_Main_Tab(2).Caption = Load_Lanuage("�༭�����ļ�", "Main", "Option_Main_Tab2", Lanuage_Now)
    .Command_Guide.Caption = Load_Lanuage("������", "Main", "Option_Main_Tab3", Lanuage_Now)
    'ѡ������
    .Label_Help_Select_Theme.Caption = Load_Lanuage("�����б�������ϵͳ���Ѿ���װ������|��ѡ������ҪӦ�õĻ��߱༭������", "Main", "Help_Select_Theme", Lanuage_Now)
    .Command_Choose_Aply_Theme.Caption = Load_Lanuage("Ӧ�õ�ϵͳ", "Main", "Command_Choose_Aply_Theme", Lanuage_Now)
    .Command_Choose_Add_Theme.Caption = Load_Lanuage("����б���û�е�����", "Main", "Command_Choose_Add_Theme", Lanuage_Now)
    .Command_Choose_Edit_Theme.Caption = Load_Lanuage("�༭������", "Main", "Command_Choose_Edit_Theme", Lanuage_Now)
    .Command_Choose_Refresh_Theme.Caption = Load_Lanuage("ˢ���б�", "Main", "Command_Choose_Refresh_Theme", Lanuage_Now)
    .Command_Down_More_Theme.Caption = Load_Lanuage("��ȡ��������", "Main", "Command_Down_More_Theme", Lanuage_Now)
    '�ֶ�Ӧ��
    .Label_mss_indro.Caption = Load_Lanuage("�Զ�Ӧ���Ӿ�����ļ����ܻ���������ʧ�ܣ��ɳ��Զ�㼸�Ρ�(�����ԱȨ������������|���һֱû��Ӧ�óɹ����������Ƿ��ƽ������⣬������ѡ����Ӿ�����ļ��Ĳ���ϵͳ�Ƿ��Ӧ", "Main", "Help_Aply_By_Hand", Lanuage_Now)
    .Command_ico_hand.Caption = Load_Lanuage("��������ͼ��", "Main", "Command_ico_hand", Lanuage_Now)
    .Command_cur_hand.Caption = Load_Lanuage("�������ָ��", "Main", "Command_cur_hand", Lanuage_Now)
    .Command_snd_hand.Caption = Load_Lanuage("����ϵͳ��Ч", "Main", "Command_snd_hand", Lanuage_Now)
    .Command_paper_hand.Caption = Load_Lanuage("���������ֽ", "Main", "Command_paper_hand", Lanuage_Now)
    .Command_window_hand.Caption = Load_Lanuage("���Ĵ�����ɫ�����", "Main", "Command_window_hand", Lanuage_Now)
    .Command_glass_hand.Caption = Load_Lanuage("����͸����ɫ", "Main", "Command_glass_hand", Lanuage_Now)
    .Command_individuation_hand.Caption = Load_Lanuage("�򿪸��Ի�", "Main", "Command_individuation_hand", Lanuage_Now)
    .Command_scr_hand.Caption = Load_Lanuage("��װ��Ļ��������", "Main", "Command_scr_hand", Lanuage_Now)
    .Command_mss_hand.Caption = Load_Lanuage("�޸��Ӿ����", "Main", "Command_mss_hand", Lanuage_Now)
    .Label_scr_hand.Caption = Load_Lanuage("��Ļ���������ļ�", "Public", "CommonDialog_Scr_Filter", Lanuage_Now)
    .Label_mss_hand.Caption = Load_Lanuage("�Ӿ�����ļ�", "Public", "CommonDialog_Mss_Filter", Lanuage_Now)
    .Command_scr_open.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
    .Command_mss_open.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
    '�༭
        '����
        .Edit_Panel_Tab(0).Caption = Load_Lanuage("������Ϣ", "Main", "Edit_Panel_Theme_info_Caption", Lanuage_Now)
        .Edit_Panel_Tab(1).Caption = Load_Lanuage("�Ӿ����", "Main", "Edit_Panel_Mss_Caption", Lanuage_Now)
        .Edit_Panel_Tab(2).Caption = Load_Lanuage("�����ֽ", "Main", "Edit_Panel_Paper_Caption", Lanuage_Now)
        .Edit_Panel_Tab(3).Caption = Load_Lanuage("����ͼ��", "Main", "Edit_Panel_Icon_Caption", Lanuage_Now)
        .Edit_Panel_Tab(4).Caption = Load_Lanuage("���ָ��", "Main", "Edit_Panel_Curson_Caption", Lanuage_Now)
        .Edit_Panel_Tab(5).Caption = Load_Lanuage("ϵͳ��Ч", "Main", "Edit_Panel_Sound_Caption", Lanuage_Now)
        .Edit_Panel_Tab(6).Caption = Load_Lanuage("��Ļ��������", "Main", "Edit_Panel_Scr_Caption", Lanuage_Now)
        
        .Edit_Panel_Frame(0).Caption = Load_Lanuage("������Ϣ", "Main", "Edit_Panel_Theme_info_Caption", Lanuage_Now)
        .Edit_Panel_Frame(1).Caption = Load_Lanuage("�Ӿ����", "Main", "Edit_Panel_Mss_Caption", Lanuage_Now)
        .Edit_Panel_Frame(2).Caption = Load_Lanuage("�����ֽ", "Main", "Edit_Panel_Paper_Caption", Lanuage_Now)
        .Edit_Panel_Frame(3).Caption = Load_Lanuage("����ͼ��", "Main", "Edit_Panel_Icon_Caption", Lanuage_Now)
        .Edit_Panel_Frame(4).Caption = Load_Lanuage("���ָ��", "Main", "Edit_Panel_Curson_Caption", Lanuage_Now)
        .Edit_Panel_Frame(5).Caption = Load_Lanuage("ϵͳ��Ч", "Main", "Edit_Panel_Sound_Caption", Lanuage_Now)
        .Edit_Panel_Frame(6).Caption = Load_Lanuage("��Ļ��������", "Main", "Edit_Panel_Scr_Caption", Lanuage_Now)
        .Command_Aply_Now.Caption = Load_Lanuage("����Ӧ��Ч��", "Main", "Command_Aply_Now", Lanuage_Now)
            '������Ϣ
            .Label_TnameC.Caption = Load_Lanuage("������ʾ����", "Main", "Label_Tname_Display", Lanuage_Now)
            .Label_TnameE.Caption = Load_Lanuage("�����ļ�����", "Main", "Label_Tname_File", Lanuage_Now)
            .Label_maker.Caption = Load_Lanuage("����������", "Main", "Label_maker", Lanuage_Now)
            .Label_maker_web.Caption = Load_Lanuage("��ַ�������ҳ", "Main", "Label_maker_web", Lanuage_Now)
            .Label_Maker_Introduce.Caption = Load_Lanuage("������Ȩ��Ϣ��˵��", "Main", "Label_Maker_Introduce", Lanuage_Now)
            .Label_Tlogo.Caption = Load_Lanuage("����LOGO:", "Main", "Label_Tlogo", Lanuage_Now)
            .Label_Logo_Preview.Caption = Load_Lanuage("Ԥ��:", "Main", "Label_Logo_Preview", Lanuage_Now)
            .Label_logo_help.Caption = Load_Lanuage("�Ƽ���������͸����PNG��ʽ|LOGO�����ʾΪ240��80���أ���256��256״̬�£�|����벻Ҫ̫��", "Main", "Label_logo_help", Lanuage_Now)
            .Command_Tlogo.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            '�Ӿ����
            .Frame_select_mss.Caption = Load_Lanuage("���ѡ��", "Main", "Frame_select_mss", Lanuage_Now)
            .mss_Aero.Caption = Load_Lanuage("Aero", "Main", "mss_Aero", Lanuage_Now)
            .Check_Alpha.Caption = Load_Lanuage("����͸��", "Main", "Check_Alpha", Lanuage_Now)
            .mss_Basic.Caption = Load_Lanuage("Basic", "Main", "mss_Basic", Lanuage_Now)
            .mss_Classic.Caption = Load_Lanuage("Classic", "Main", "mss_Classic", Lanuage_Now)
            .Command_getcolor.Caption = Load_Lanuage("һ��ȡɫ", "Main", "Command_getcolor", Lanuage_Now)
            .System_Color_Tab(0).Caption = Load_Lanuage("���ӻ�������", "Main", "System_Color_Tab1", Lanuage_Now)
            .System_Color_Frame(0).Caption = .System_Color_Tab(0).Caption
            .System_Color_Tab(1).Caption = Load_Lanuage("������ɫ�����", "Main", "System_Color_Tab2", Lanuage_Now)
            .System_Color_Frame(1).Caption = .System_Color_Tab(1).Caption
            .Label_mss.Caption = Load_Lanuage("�Ӿ�����ļ�", "Main", "Label_mss", Lanuage_Now)
            .Command_mss.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            .Label_ColorizationColor.Caption = Load_Lanuage("����ɫ", "Main", "Label_ColorizationColor", Lanuage_Now)
            .Label_ColorizationColor_alpha.Caption = Load_Lanuage("����ɫ͸����", "Main", "Label_ColorizationColor_alpha", Lanuage_Now)
            .Label_ColorizationColorBalance.Caption = Load_Lanuage("����ɫƽ��", "Main", "Label_ColorizationColorBalance", Lanuage_Now)
            .Label_ColorizationAfterglow.Caption = Load_Lanuage("������ɫ", "Main", "Label_ColorizationAfterglow", Lanuage_Now)
            .Label_ColorizationAfterglow_alpha.Caption = Load_Lanuage("������ɫ͸����", "Main", "Label_ColorizationAfterglow_alpha", Lanuage_Now)
            .Label_ColorizationAfterglowBalance.Caption = Load_Lanuage("������ɫƽ��", "Main", "Label_ColorizationAfterglowBalance", Lanuage_Now)
            .Label_ColorizationGlassReflectionIntensity.Caption = Load_Lanuage("Aero��������", "Main", "Label_ColorizationGlassReflectionIntensity", Lanuage_Now)
            .Label_ColorizationBlurBalance.Caption = Load_Lanuage("ģ��ƽ��", "Main", "Label_ColorizationBlurBalance", Lanuage_Now)
            .Color_Warn.Caption = Load_Lanuage("�Լ��༭��ɫ���ܵ���һЩ��ֵ���ɫ����ʹ��һ��ȡɫ����", "Main", "Color_Warn", Lanuage_Now)
            .Label_Classic_Style.Caption = Load_Lanuage("������Ԥ��: ", "Main", "Label_Classic_Style", Lanuage_Now)
            .Check_insert_system_color.Caption = Load_Lanuage("���Զ�����ɫ���뵽����������BAT�ļ��С�����ѡ��Ϊ�÷��ϵͳĬ��ֵ����", "Main", "Check_insert_system_color", Lanuage_Now)
            '��ֽ
            .Label_paper_index.Caption = Load_Lanuage("����ֽ�ļ�:", "Main", "Label_paper_index", Lanuage_Now)
            .Command_paper.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            .Label_paper_style.Caption = Load_Lanuage("��ֽ��ʾģʽ:", "Main", "Label_paper_style", Lanuage_Now)
            .Label_paper_change_time.Caption = Load_Lanuage("�õ�Ƭ�л�ʱ��:", "Main", "Label_paper_change_time", Lanuage_Now)
            .Check_paper_change.Caption = Load_Lanuage("�����л�", "Main", "Check_paper_change", Lanuage_Now)
            .Label_paper_files.Caption = Load_Lanuage("��ֽ�õ�Ƭ�ļ���:", "Main", "Label_paper_files", Lanuage_Now)
            .Command_paper_files.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            .Papers_Edit_Allow.Caption = Load_Lanuage("����༭ͼƬ�б�", "Main", "Papers_Edit_Allow", Lanuage_Now)
            .Papers_Edit_Select_All.Caption = Load_Lanuage("ѡ��ȫ��", "Main", "Papers_Edit_Select_All", Lanuage_Now)
            .Papers_Edit_Clear.Caption = Load_Lanuage("ȫ����ѡ", "Main", "Papers_Edit_Clear", Lanuage_Now)
            'ͼ��
            For i = 0 To 5
                .Command_icon(i).Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            Next
            '���ָ��
            .Label_cur.Caption = Load_Lanuage("��������", "Main", "Label_cur", Lanuage_Now)
            .Command_cur.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            .cur_default.Caption = Load_Lanuage("ʹ��Ĭ��ֵ(Win7)", "Main", "cur_default", Lanuage_Now)
            '��Ч
            .Check_snd.Caption = Load_Lanuage("ʹ���Ѵ��ڵķ���", "Main", "Check_snd", Lanuage_Now)
            .Labe_sound_name_C.Caption = Load_Lanuage("������������", "Main", "Labe_sound_name_C", Lanuage_Now)
            .Label_sound_name_E.Caption = Load_Lanuage("��������Ӣ�ļ�д", "Main", "Label_sound_name_E", Lanuage_Now)
            .sound_Play.Caption = Load_Lanuage("����", "Main", "sound_Play", Lanuage_Now)
            .sound_Stop.Caption = Load_Lanuage("ֹͣ", "Main", "sound_Stop", Lanuage_Now)
            .Command_sound.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
            '��Ļ��������
            .Label_scr_url.Caption = Load_Lanuage("�����ļ������������������ռ��ɣ�", "Main", "Label_scr_url", Lanuage_Now)
            .Label_scr_wait.Caption = Load_Lanuage("�ȴ�:", "Main", "Label_scr_wait", Lanuage_Now)
            .Label_scr_wait_min.Caption = Load_Lanuage("����", "Main", "Label_scr_wait_min", Lanuage_Now)
            .Command_scr.Caption = Load_Lanuage("���", "Public", "Command_Select", Lanuage_Now)
End With
End Sub
