VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ������"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5865
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame_AutoPaper 
      Caption         =   "���ͱ�ֽ�б����Զ�������ֽ������"
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   5655
      Begin VB.OptionButton Option_AutoPaper_A 
         Caption         =   "ѯ��"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option_AutoPaper_Y 
         Caption         =   "��"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option_AutoPaper_N 
         Caption         =   "��"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check_frmLoad 
      Caption         =   "��������ʱ��������������"
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Value_Aplha_Back_Color 
      Height          =   270
      Left            =   720
      MaxLength       =   10
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
   End
   Begin VB.PictureBox Show_Aplha_Back_Color 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame_Soft_Glass 
      Caption         =   "��������ʾ���"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   5655
      Begin VB.OptionButton Aero_Glass 
         Caption         =   "Aeroȫ����"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton Aero_Normal 
         Caption         =   "��ͨ"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.ComboBox Combo_SystemRoot 
      Height          =   300
      ItemData        =   "Options.frx":0000
      Left            =   120
      List            =   "Options.frx":000A
      TabIndex        =   14
      Text            =   "����ϵͳ����λ��"
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Command_Find_Lanuages 
      Caption         =   "��ȡ����Find More"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command_Done 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command_Aply 
      Caption         =   "Aply"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.ComboBox Combo_Lanuage 
      Height          =   300
      ItemData        =   "Options.frx":0022
      Left            =   120
      List            =   "Options.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   4575
   End
   Begin VB.Frame Frame_SystemTextShow 
      Caption         =   "����ϵͳ����"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      Begin VB.ComboBox Combo_Snd_Style 
         Height          =   300
         ItemData        =   "Options.frx":0026
         Left            =   2760
         List            =   "Options.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton SystemTextShow_ini 
         Caption         =   "�������ļ���ȡ"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton SystemTextShow_Sys 
         Caption         =   "��ϵͳ��ȡ"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.Label Label_Snd_Style 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ч�б�汾"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.ComboBox Combo_SysPath 
      Height          =   300
      ItemData        =   "Options.frx":002A
      Left            =   3360
      List            =   "Options.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label_Aplha_Back_Color 
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫԤ��������ɫ"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      ToolTipText     =   "Aero��ɫԤ��͸���ȵı�����ɫ"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label_Lanuage 
      BackStyle       =   0  'Transparent
      Caption         =   "�������/Lanuages"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label_SysPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�����ɺ��ֻ�������"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "ϵͳ��·���Զ�ת���ɻ�������"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label_SystemRoot 
      BackStyle       =   0  'Transparent
      Caption         =   "����ϵͳ����λ��"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "�����ϵͳ���ԣ��༭�ǵ�ǰϵͳ������ʱ���޸�"
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo_Lanuage_Click()
If Combo_Lanuage.ListCount > 0 Then
Dim i As Integer, x As Integer, sNum As Integer
    x = Combo_Lanuage.ListIndex
    Combo_Snd_Style.Clear '������������б�
    Combo_Snd_Style.AddItem Load_Lanuage("��ѡ��", "OptionsForm", "Combo_Snd_Style", Lanuage_Now), 0
    If Combo_Lanuage.ListIndex >= 1 Then 'And Combo_Lanuage.ListIndex <= UBound(Lanuages)
        For i = 1 To 32767
            If GetFromIni("Sounds", "List" & i, Lanuages(x)) <> "" Then
                Combo_Snd_Style.AddItem GetFromIni("Sounds", "List" & i, Lanuages(x)), i
                sNum = sNum + 1
            Else
                If sNum >= 1 And Val(GetFromIni("Option", "SoundList", Config_Url)) <= sNum Then
                    Combo_Snd_Style.ListIndex = Val(GetFromIni("Option", "SoundList", Config_Url))
                ElseIf sNum >= 1 Then
                    Combo_Snd_Style.ListIndex = 1
                Else
                    Combo_Snd_Style.ListIndex = 0
                End If
                Exit For
            End If
        Next i
    Else
        Combo_Snd_Style.ListIndex = 0
    End If
End If
End Sub

Private Sub Combo_SystemRoot_Click()
If Combo_SystemRoot.ListIndex = 1 Then
    Dim url_temp As String
    url_temp = Environ("SystemRoot")
    url_temp = BrowseForFolderByPath(url_temp, Load_Lanuage("��ѡ�����ϵͳ�����ļ���", "OptionsForm", "Choose_Path", Lanuage_Now), Me)
    If url_temp <> "" Then
        Combo_SystemRoot.text = url_temp
    End If
End If
End Sub

'Ӧ�ð�ť
Private Sub Command_Aply_Click()
Dim Lanuage_Old As String
Lanuage_Old = Lanuage_Now
    If Combo_Lanuage.ListIndex <= 0 Then '����
        Lanuage_Now = 0 '��ǰϵͳ����
        If Lanuage_Old <> Lanuage_Now Then Call Change_Lanuage(Lanuage_Now)   '����ı������Ծ��л�����
        Call WriteIni("Option", "Lanuage", 0, Config_Url)
    ElseIf GetFromIni("info", "ShortName", Lanuages(Combo_Lanuage.ListIndex)) <> "" Then
        Lanuage_Now = Combo_Lanuage.ListIndex '��ǰϵͳ����
        If Lanuage_Old <> Lanuage_Now Then Call Change_Lanuage(Lanuage_Now)   '����ı������Ծ��л�����
        Call WriteIni("Option", "Lanuage", GetFromIni("info", "ShortName", Lanuages(Combo_Lanuage.ListIndex)), Config_Url)
    Else
        Lanuage_Now = 0 '��ǰϵͳ����
        If Lanuage_Old <> Lanuage_Now Then Call Change_Lanuage(Lanuage_Now)   '����ı������Ծ��л�����
        Call WriteIni("Option", "Lanuage", 0, Config_Url)
    End If
    '�����б�
    If SystemTextShow_Sys.value = True Then
        SystemTextShow = False
        Call WriteIni("Option", "SystemTextShow", 0, Config_Url)
    ElseIf SystemTextShow_ini.value = True Then
        SystemTextShow = True
        Call WriteIni("Option", "SystemTextShow", 1, Config_Url)
    End If
    Call WriteIni("Option", "SoundList", Combo_Snd_Style.ListIndex, Config_Url)
    Sound_Style = Combo_Snd_Style.ListIndex
    'ȫ����
    If Aero_Normal.value = True Then
        Call WriteIni("Option", "Aero", 0, Config_Url)
    ElseIf Aero_Glass.value = True Then
        Call WriteIni("Option", "Aero", 1, Config_Url)
    End If
    
    '�Զ���ֽ
    If Option_AutoPaper_N.value = True Then
        Call WriteIni("Option", "AutoPaper", 0, Config_Url)
    ElseIf Option_AutoPaper_Y.value = True Then
        Call WriteIni("Option", "AutoPaper", 1, Config_Url)
    Else
        Call WriteIni("Option", "AutoPaper", 2, Config_Url)
    End If

    '����ϵͳ����λ��
    If Combo_SystemRoot.ListIndex = 0 Then
        Call WriteIni("Option", "SystemRoot", 0, Config_Url)
        SysRoot = 0
    Else
        Call WriteIni("Option", "SystemRoot", Combo_SystemRoot.text, Config_Url)
        SysRoot = Combo_SystemRoot.text
    End If
    '��������
    Call WriteIni("Option", "SysPath_Default", Combo_SysPath.ListIndex, Config_Url)
    SysPath = Combo_SysPath.ListIndex
    '͸����ɫ
    Call WriteIni("Option", "Aplha_Back_Color", Value_Aplha_Back_Color.text, Config_Url)
    Aplha_Back_Color = x16_to_x10(Value_Aplha_Back_Color.text)
    If Main.Visible = False Then
        '��һ����Ϊ������ɫ����change����������ˢ����ɫ
        Dim Color_Temp1 As String, Color_Temp2 As String
        Color_Temp1 = Main.Value_ColorizationColor
        Color_Temp2 = Main.Value_ColorizationAfterglow
        Main.Value_ColorizationColor = " "
        Main.Value_ColorizationAfterglow = " "
        Main.Value_ColorizationColor = Color_Temp1
        Main.Value_ColorizationAfterglow = Color_Temp2
    End If
    'ϵͳ����
    Call WriteIni("Option", "Load_Guide", Check_frmLoad.value, Config_Url)
End Sub
'ȡ����ť
Private Sub Command_Cancel_Click()
    Unload Me
End Sub
'ȷ����ť
Private Sub Command_Done_Click()
    Call Command_Aply_Click
    Call Get_Options '�ٴ����¶�ȡ����
    If Main.Visible = False Then frmLoad.Show '���û�����������������������
    Call Command_Cancel_Click
    MsgBox Load_Lanuage("�������ÿ�����Ҫ���������������Ч", "OptionsForm", "Awoke_Reset", Lanuage_Now)
    Call Creat_Default
    
Dim glass_ok_load As Boolean
If GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive") = 0 Or GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "Composition") = 0 Then
Else
    If Main.Visible = True Then '����ҳ����ʾʱ
        If GetFromIni("Option", "Aero", Config_Url) <> "" Then
            glass_ok_load = GetFromIni("Option", "Aero", Config_Url)
        Else
            glass_ok_load = 0
        End If
        
        If glass_ok_load <> glass_ok Then
            Exit_ok = False
            Unload Main
            Main.Show
        End If
    End If
End If
End Sub

Private Sub Command_Find_Lanuages_Click()
        Call ShellExecute(Me.hwnd, vbNullString, "http://www.mapaler.com/mapletheme/lanuages", vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub Form_Initialize()
    Set Lanuages = New Collection
    'Erase Lanuages '���������
    Call GetFileName(url_to_N(App.Path & "\Lanuages"), "ini", Lanuages) '��ȡ�ļ��б�
End Sub

Private Sub Form_Load()
    Call Change_Lanuage(Lanuage_Now)

    'ȫ������

If glass_ok = True Then
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributesByColor Me.hwnd, m_transparencyKey, 0, LWA_COLORKEY
    On Error GoTo ern
    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    'MsgBox "1"
    DwmIsCompositionEnabled en
    If en Then
        'MsgBox "2"
        DwmExtendFrameIntoClientArea Me.hwnd, mg
        'MsgBox "OK!"
    End If
    GoTo Next_Glass
ern:
    MsgBox Err.description
Next_Glass:
End If
'ȫ������
If glass_ok = True Then
    Frame_SystemTextShow.BackColor = m_transparencyKey
    Frame_Soft_Glass.BackColor = m_transparencyKey
    Frame_AutoPaper.BackColor = m_transparencyKey
    Option_AutoPaper_Y.BackColor = m_transparencyKey
    Option_AutoPaper_N.BackColor = m_transparencyKey
    Option_AutoPaper_A.BackColor = m_transparencyKey
    Check_frmLoad.BackColor = m_transparencyKey
    SystemTextShow_Sys.BackColor = m_transparencyKey
    SystemTextShow_ini.BackColor = m_transparencyKey
    Aero_Normal.BackColor = m_transparencyKey
    Aero_Glass.BackColor = m_transparencyKey
Else
    Frame_SystemTextShow.BackColor = &H8000000F
    Frame_Soft_Glass.BackColor = &H8000000F
    Frame_AutoPaper.BackColor = &H8000000F
    Option_AutoPaper_Y.BackColor = &H8000000F
    Option_AutoPaper_N.BackColor = &H8000000F
    Option_AutoPaper_A.BackColor = &H8000000F
    Check_frmLoad.BackColor = &H8000000F
    SystemTextShow_Sys.BackColor = &H8000000F
    SystemTextShow_ini.BackColor = &H8000000F
    Aero_Normal.BackColor = &H8000000F
    Aero_Glass.BackColor = &H8000000F
End If

Dim i As Integer, j As Integer
Dim x As Integer
Dim Lanuage_ShortName As String
    Me.Icon = Main.Icon 'ͼ�걣�ֺ�������һ��
    Command_Find_Lanuages.Caption = "��ȡ����" & vbCrLf & "Find More"
    
'��������ļ��б�
    Combo_Lanuage.Clear '��������б�
    Combo_Snd_Style.Clear '������������б�
'On Error GoTo ErrHandler
    Combo_Lanuage.AddItem Load_Lanuage("��ʹ������/Don't use lanuage file", "OptionsForm", "Combo_Lanuage", Lanuage_Now), 0
    Combo_Snd_Style.AddItem Load_Lanuage("��ѡ��", "OptionsForm", "Combo_Snd_Style", Lanuage_Now), 0
    'If UBound(Lanuages) > 0 Then '����������ļ�
    If Lanuages.count > 0 Then '����������ļ�
        'For i = 1 To UBound(Lanuages)
        For i = 1 To Lanuages.count
            Combo_Lanuage.AddItem GetFromIni("info", "DisplayName", Lanuages(i)), i
        Next
        Lanuage_ShortName = GetFromIni("Option", "Lanuage", Config_Url) '��ȡ����
        'For i = 1 To UBound(Lanuages)
        For i = 1 To Lanuages.count
            If Lanuage_ShortName = GetFromIni("info", "ShortName", Lanuages(i)) Then
                Combo_Lanuage.ListIndex = i
                Lanuage_Now = i '���õ�ǰ����
                Exit For
            Else
                Combo_Lanuage.ListIndex = 0
            End If
        Next i
    Else
        Lanuage_Now = 0
        x = MsgBox("û�м�⵽�����ļ����Ƿ�ǰ�������վ�������ԡ�" + vbLf + "No found lanuage.Do you want to download lanuages from the soft website?", 36, "No found Options and Lanuages")
        If x = 6 Then '��
            Call ShellExecute(Me.hwnd, vbNullString, "http://www.mapaler.com/mapletheme/lanuages.html", vbNullString, vbNullString, SW_SHOWNORMAL)
        Else
        End If
        Combo_Lanuage.ListIndex = 0
        MsgBox "a"
    End If
'GoTo Next1
'ErrHandler: '�±����û�������ļ���
'        Lanuage_Now = 0
'    x = MsgBox("û�м�⵽�����ļ����Ƿ�ǰ�������վ�������ԡ�" + vbLf + "No found lanuage.Do you want to download lanuages from the soft website?", 36, "No found Options and Lanuages")
'    If x = 6 Then '��
'        Call ShellExecute(Me.hwnd, vbNullString, "http://www.mapaler.com/mapletheme/lanuages.html", vbNullString, vbNullString, SW_SHOWNORMAL)
'    Else
'    End If
'    Combo_Lanuage.ListIndex = 0
'Next1:

Combo_SystemRoot.list(0) = Load_Lanuage("��ϵͳ��ȡ", "OptionsForm", "Combo_SystemRoot1", Lanuage_Now)
Combo_SystemRoot.list(1) = Load_Lanuage("�Զ���", "OptionsForm", "Combo_SystemRoot2", Lanuage_Now)
'����ϵͳ�ļ���
If GetFromIni("Option", "SystemRoot", Config_Url) <> "" Then
    If GetFromIni("Option", "SystemRoot", Config_Url) <> "0" Then
        Combo_SystemRoot = GetFromIni("Option", "SystemRoot", Config_Url)
    Else 'Ϊ��
        Combo_SystemRoot.ListIndex = 0
        Combo_SystemRoot.text = Load_Lanuage("��ϵͳ��ȡ", "OptionsForm", "Combo_SystemRoot1", Lanuage_Now)
    End If
Else 'Ϊ��
    Combo_SystemRoot.ListIndex = 0
    Combo_SystemRoot.text = Load_Lanuage("��ϵͳ��ȡ", "OptionsForm", "Combo_SystemRoot1", Lanuage_Now)
End If
'��������
If GetFromIni("Option", "SysPath_Default", Config_Url) <> "" Then
    If GetFromIni("Option", "SysPath_Default", Config_Url) >= "0" And GetFromIni("Option", "SysPath_Default", Config_Url) < "3" Then
        Combo_SysPath.ListIndex = GetFromIni("Option", "SysPath_Default", Config_Url)
    Else
        Combo_SysPath.ListIndex = 0
    End If
Else
    Combo_SysPath.ListIndex = 0
End If
Combo_SysPath.list(2) = Load_Lanuage("��ת����������", "OptionsForm", "Combo_SysPath", Lanuage_Now)
'�����б�
If GetFromIni("Option", "SystemTextShow", Config_Url) <> "" Then
    If GetFromIni("Option", "SystemTextShow", Config_Url) = "0" Then
        SystemTextShow_Sys.value = True
        Label_Snd_Style.Enabled = False
        Combo_Snd_Style.Enabled = False
    Else
        SystemTextShow_ini.value = True
    End If
Else
    SystemTextShow_Sys.value = True
    Label_Snd_Style.Enabled = False
    Combo_Snd_Style.Enabled = False
End If
'ȫ����
If GetFromIni("Option", "Aero", Config_Url) <> "" Then
    If GetFromIni("Option", "Aero", Config_Url) = "0" Then
        Aero_Normal.value = True
    Else
        Aero_Glass.value = True
    End If
Else
    Aero_Normal.value = True
End If
'�Զ���ֽ
If GetFromIni("Option", "AutoPaper", Config_Url) <> "" Then
    If GetFromIni("Option", "AutoPaper", Config_Url) = "0" Then
        Option_AutoPaper_N.value = True
    ElseIf GetFromIni("Option", "AutoPaper", Config_Url) = "1" Then
        Option_AutoPaper_Y.value = True
    Else
        Option_AutoPaper_A.value = True
    End If
Else
    Option_AutoPaper_A.value = True
End If
'ȫ����
If GetFromIni("Option", "Aplha_Back_Color", Config_Url) <> "" Then
    Value_Aplha_Back_Color.text = GetFromIni("Option", "Aplha_Back_Color", Config_Url)
Else
    Value_Aplha_Back_Color.text = "FFFFFF"
End If
'��������
If GetFromIni("Option", "Load_Guide", Config_Url) <> "" Then
    Check_frmLoad.value = GetFromIni("Option", "Load_Guide", Config_Url)
Else
    Check_frmLoad.value = 0
End If

Dim LanuageVer() As String, LanuageNeed() As String, NeedNew As Boolean
LanuageVer = Split(Load_Lanuage("1.0.0", "info", "AppVersion", Lanuage_Now), ".")
LanuageNeed = Split(Lanuage_Need, ".")
NeedNew = False
For i = 0 To UBound(LanuageVer)
    If LanuageNeed(i) > LanuageVer(i) Then
        NeedNew = True
        Exit For
    End If
Next
If NeedNew Then
    Call MsgBox(Load_Lanuage("Your lanuage is old,,can't show all part of the soft,suggest you find newest lanuage", "info", "Warn", Lanuage_Now), 64)
End If
End Sub

Private Sub Form_Paint()
'ȫ������
If glass_ok = True Then
    Dim hBrush As Long, m_Rect As rect, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hDC, hBrush)
    GetClientRect Me.hwnd, m_Rect

    FillRect Me.hDC, m_Rect, hBrush
    SelectObject Me.hDC, hBrushOld

    DeleteObject hBrush
End If
'ȫ����
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Main.Visible = False And frmLoad.Visible = False Then End '���û������ĳ�����������˳�����
End Sub

Private Sub Show_Aplha_Back_Color_Click()
Dim i%
Dim Color_Temp As Long
Rem ��ȡ��ɫ
On Error GoTo ErrHandler
    Main.CommonDialog1.ShowColor
    Show_Aplha_Back_Color.BackColor = Main.CommonDialog1.Color
    Color_Temp = RGB_To_BGR(Main.CommonDialog1.Color)
    Value_Aplha_Back_Color = x10_to_x16(Color_Temp, 6)
Exit Sub
ErrHandler:
'�û�����ȡ������ť��
Exit Sub
End Sub

Private Sub SystemTextShow_ini_Click()
    Label_Snd_Style.Enabled = True
    Combo_Snd_Style.Enabled = True
End Sub

Private Sub SystemTextShow_Sys_Click()
    Label_Snd_Style.Enabled = False
    Combo_Snd_Style.Enabled = False
End Sub

Private Sub Value_Aplha_Back_Color_Change()
Dim Color_Temp As Long
Color_Temp = x16_to_x10(text_to_color(Value_Aplha_Back_Color))
Show_Aplha_Back_Color.BackColor = RGB_To_BGR(Color_Temp)
End Sub
