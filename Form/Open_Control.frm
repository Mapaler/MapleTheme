VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Open_Control 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ֶ��޸�"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Open_Control.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   6495
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command_exit 
      Cancel          =   -1  'True
      Caption         =   "�˳�"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command_mss_open 
      Caption         =   "���"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command_scr_open 
      Caption         =   "���"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Url_scr_hand 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Url_mss_hand 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command_mss_hand 
      Caption         =   "�޸��Ӿ����"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command_cur_hand 
      Caption         =   "�޸������"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   680
      Width           =   1935
   End
   Begin VB.CommandButton Command_snd_hand 
      Caption         =   "�޸�ϵͳ��Ч"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1240
      Width           =   1935
   End
   Begin VB.CommandButton Command_paper_hand 
      Caption         =   "�޸ı�ֽ"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command_ico_hand 
      Caption         =   "�޸�ͼ��"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command_scr_hand 
      Caption         =   "��װ��Ļ��������"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command_glass_hand 
      Caption         =   "�޸�͸����ɫ"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command_window_hand 
      Caption         =   "�޸Ĵ�����ɫ�����"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command_individuation_hand 
      Caption         =   "�򿪸��Ի�"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label_mss_indro 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   5535
   End
End
Attribute VB_Name = "Open_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_cur_Click()
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0,1") '�޸����ָ��
End Sub

Private Sub Command_exit_Click()
    Unload Me
End Sub

Private Sub Command_glass_Click()
If System_Ver < 6 Then
    MsgBox "��⵽����ϵͳΪ" & strOSversion & "�����Ĳ���ϵͳ�汾���޴˹��ܡ�"
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageColorization") '�޸�͸����ɫ
End If
End Sub

Private Sub Command_ico_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '�޸�ͼ��
    MsgBox "��⵽����ϵͳΪ" & strOSversion & "���޸�ͼ�������Զ������水ť��"
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '�޸�ͼ��
    MsgBox "��⵽����ϵͳΪ" & strOSversion & "����ϵͳ�ҾͲ�֪����ʲô�����ˡ���"
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '�޸�ͼ��
End If
End Sub

Private Sub Command_individuation_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,-1") 'XP������
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '�򿪸��Ի�
End If
End Sub

Private Sub Command_mss_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = "ѡ�� Windows �Ӿ���ʽ�ļ�"
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = "Windows �Ӿ���ʽ�ļ� (*.msstyles)|*.msstyles"
    If url_mss <> "" Then '��Ϊ��
        CommonDialog1.filename = url_to_N(Url_mss_hand) '��ʱĬ��ѡ��ǰ�ļ�
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    Url_mss_hand = CommonDialog1.filename
Exit Sub
ErrHandler:
'�û�����ȡ������ť��
Exit Sub
End Sub

Private Sub Command_paper_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") 'XP������ֽ
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageWallpaper") 'Win7������ֽ
End If
End Sub

Private Sub Command_scr_Click()
Dim x As Integer
If url_scr <> "" And Dir(url_to_N(url_scr)) <> "" Then '��Ϊ�����ļ�����
    Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '��Ļ��������
Else
    x = MsgBox("���������Ļ���������ļ���ַ����ⲻ���ڣ��Ƿ���Ȼ����Ӧ�õ�ϵͳ��", 4, "�ļ�������")
    If x = 6 Then '��
        Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '��Ļ��������
    End If
End If
End Sub

Private Sub Command_scr_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = "ѡ����Ļ���������ļ�"
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = "��Ļ�������� (*.scr)|*.scr"
    If url_scr <> "" Then '��Ϊ��
        CommonDialog1.filename = url_to_N(Url_scr_hand) '��ʱĬ��ѡ��ǰ�ļ�
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    Url_scr_hand = CommonDialog1.filename
Exit Sub
ErrHandler:
'�û�����ȡ������ť��
Exit Sub
End Sub

Private Sub Command_snd_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,1") 'XP�޸�ϵͳ��Ч
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,2") 'Win7�޸�ϵͳ��Ч
End If
End Sub

Private Sub Command_window_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '�򿪸��Ի�
    MsgBox "��⵽����ϵͳΪ" & strOSversion & "���޸Ĵ�����ɫ�����߼���ť��"
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '�򿪸��Ի�
    MsgBox "��⵽����ϵͳΪ" & strOSversion & "����ϵͳ�ҾͲ�֪����ʲô�����ˡ���"
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,advanced,@advanced") '�򿪴������ú����
End If
End Sub

Private Sub Command_mss_Click()
Dim x As Integer
If url_mss <> "" And Dir(url_to_N(url_mss)) <> "" Then '��Ϊ�����ļ�����
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
    Call Shell("net stop Themes")
    Call Shell("net start Themes")
Else
    x = MsgBox("��������Ӿ�����ļ���ַ����ⲻ���ڣ��Ƿ���Ȼ����Ӧ�õ�ϵͳ��", 4, "�ļ�������")
    If x = 6 Then '��
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
        Call Shell("net stop Themes")
        Call Shell("net start Themes")
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Icon = Main.Icon
    Label_mss_indro.Caption = "�Զ�Ӧ���Ӿ�����ļ����ܻ���������ʧ�ܣ��ɳ��Զ�㼸�Ρ�(�����ԱȨ������������)" & vbCrLf & "���һֱû��Ӧ�óɹ����������Ƿ��ƽ������⣬������ѡ����Ӿ�����ļ��Ĳ���ϵͳ�Ƿ��Ӧ"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub
