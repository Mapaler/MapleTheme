VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ����������"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3015
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox Check_frmLoad 
      Caption         =   "�´β��ٳ��ֱ�����"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Frame Frame_Edit 
      Caption         =   "�༭/����Windows����"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
      Begin VB.CommandButton Command_Edit 
         Caption         =   "�򿪱༭��"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame_Basic 
      Caption         =   "Window7��ͥ��ͨ��Ӧ������"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command_theme_to_Bat 
         Caption         =   "�Զ�Ӧ�����⵽ϵͳ"
         Default         =   -1  'True
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command_Open_Control 
         Caption         =   "�ֶ�Ӧ������"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_frmLoad_Click()
    Call WriteIni("Option", "Load_Guide", Check_frmLoad.value, Config_Url)
End Sub

Private Sub Command_Edit_Click()
Main.Option_Main_Tab(2).value = True
Main.Show
Me.Hide
End Sub

Private Sub Command_Open_Control_Click()
Main.Option_Main_Tab(1).value = True
Main.Show
Me.Hide
End Sub

Private Sub Command_theme_to_Bat_Click()

Main.Option_Main_Tab(0).value = True
Main.Show
Me.Hide
'Dim save_url As String
''ֱ�Ӵ򿪺ͱ���
'On Error GoTo ErrHandler
'    Main.CommonDialog1.DialogTitle = "���Ѵ��ڵ������ļ�"
'    Main.CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
'    Main.CommonDialog1.Filter = "�����ļ� (*.theme)|*.theme"
'    Main.CommonDialog1.ShowOpen
'    Call Load_theme(Main.CommonDialog1.filename)  '���ö�ȡ����
'
'    Main.CommonDialog1.Flags = cdlOFNOverwritePrompt
'    Main.CommonDialog1.DialogTitle = "����Win7��ͥ��Ӧ�������BAT�ļ�"
'    Main.CommonDialog1.Filter = "�������ļ� (*.bat)|*.bat"
'    If Main.T_name_C <> "" Then '��Ϊ��
'        Main.CommonDialog1.filename = "Ӧ�� " & Main.T_name_C.text & " ��ϵͳ"  '��ʱĬ��ѡ��ǰ�ļ�
'    Else
'        Main.CommonDialog1.filename = "Ӧ���ҵ�����"
'    End If
'    Main.CommonDialog1.ShowSave
'    save_url = Main.CommonDialog1.filename
'    Save_Bat (save_url) '����BAT���ɳ���
'
'Main.CommonDialog1.filename = ""
'Exit Sub
'ErrHandler:
''�û�����ȡ������ť��
End Sub

Private Sub Form_Initialize()
    Config_Url = App.Path & "\config.ini"  '�����ļ�·��
    Call Get_Ststem_Ver '���ϵͳ�汾
    Call Get_Options '��ȡ����
    
    Call Change_Lanuage(Lanuage_Now)
    Auto_Update = True '������Զ�����
    Main.Timer_Update.Enabled = True '���ü�����ʱ��ؼ�Ϊ����
    '���û�п�AeroЧ���ͽ�ֹ��Aero����
    
'��ȡ��������
If GetFromIni("Option", "Load_Guide", Config_Url) <> "" Then
    Check_frmLoad.value = GetFromIni("Option", "Load_Guide", Config_Url)
    If Check_frmLoad.value = 1 Then '�������ʱ����ʾ�˴���Ϊѡ�У���ֱ�Ӵ�������
      '������ѡ����һ���������
        If GetFromIni("Option", "Main_Tab", Config_Url) <> "" Then
            Main.Option_Main_Tab(GetFromIni("Option", "Main_Tab", Config_Url)).value = True
        Else
            Main.Option_Main_Tab(0).value = True
        End If
        Main.Show
        Me.Hide
    End If
Else
    Check_frmLoad.value = 0
End If
End Sub

Private Sub Form_Load()
'    Me.Caption = "������� - ѡ����������"
    Me.Icon = Main.Icon
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
    Frame_Basic.BackColor = m_transparencyKey
    Frame_Edit.BackColor = m_transparencyKey
    Check_frmLoad.BackColor = m_transparencyKey
Else
    Frame_Basic.BackColor = &H8000000F
    Frame_Edit.BackColor = &H8000000F
    Check_frmLoad.BackColor = &H8000000F
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
    End
End Sub
