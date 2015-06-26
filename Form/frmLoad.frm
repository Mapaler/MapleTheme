VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择启动任务"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3015
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check_frmLoad 
      Caption         =   "下次不再出现本界面"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Frame Frame_Edit 
      Caption         =   "编辑/生成Windows主题"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
      Begin VB.CommandButton Command_Edit 
         Caption         =   "打开编辑器"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame_Basic 
      Caption         =   "Window7家庭普通版应用主题"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command_theme_to_Bat 
         Caption         =   "自动应用主题到系统"
         Default         =   -1  'True
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command_Open_Control 
         Caption         =   "手动应用主题"
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
''直接打开和保存
'On Error GoTo ErrHandler
'    Main.CommonDialog1.DialogTitle = "打开已存在的主题文件"
'    Main.CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
'    Main.CommonDialog1.Filter = "主题文件 (*.theme)|*.theme"
'    Main.CommonDialog1.ShowOpen
'    Call Load_theme(Main.CommonDialog1.filename)  '调用读取主题
'
'    Main.CommonDialog1.Flags = cdlOFNOverwritePrompt
'    Main.CommonDialog1.DialogTitle = "保存Win7家庭版应用主题的BAT文件"
'    Main.CommonDialog1.Filter = "批处理文件 (*.bat)|*.bat"
'    If Main.T_name_C <> "" Then '不为空
'        Main.CommonDialog1.filename = "应用 " & Main.T_name_C.text & " 到系统"  '打开时默认选择当前文件
'    Else
'        Main.CommonDialog1.filename = "应用我的主题"
'    End If
'    Main.CommonDialog1.ShowSave
'    save_url = Main.CommonDialog1.filename
'    Save_Bat (save_url) '调用BAT生成程序
'
'Main.CommonDialog1.filename = ""
'Exit Sub
'ErrHandler:
''用户按“取消”按钮。
End Sub

Private Sub Form_Initialize()
    Config_Url = App.Path & "\config.ini"  '设置文件路径
    Call Get_Ststem_Ver '检测系统版本
    Call Get_Options '读取设置
    
    Call Change_Lanuage(Lanuage_Now)
    Auto_Update = True '这次是自动更新
    Main.Timer_Update.Enabled = True '设置检测更新时间控件为可用
    '如果没有开Aero效果就禁止以Aero启动
    
'读取引导设置
If GetFromIni("Option", "Load_Guide", Config_Url) <> "" Then
    Check_frmLoad.value = GetFromIni("Option", "Load_Guide", Config_Url)
    If Check_frmLoad.value = 1 Then '如果启动时不显示此窗口为选中，则直接打开主窗口
      '主界面选择哪一个面板启动
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
'    Me.Caption = "枫谷主题 - 选择启动任务"
    Me.Icon = Main.Icon
    '全玻璃↓

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
'全玻璃↑
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
'全玻璃↓
If glass_ok = True Then
    Dim hBrush As Long, m_Rect As rect, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hDC, hBrush)
    GetClientRect Me.hwnd, m_Rect

    FillRect Me.hDC, m_Rect, hBrush
    SelectObject Me.hDC, hBrushOld

    DeleteObject hBrush
End If
'全玻璃

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
