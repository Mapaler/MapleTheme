VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CreatGuide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "主题生成向导"
   ClientHeight    =   5880
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   12270
   Icon            =   "CreatGuide.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   12270
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   1
      Left            =   2040
      ScaleHeight     =   4395
      ScaleWidth      =   11205
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   11205
      Begin VB.Frame Frame_BT_Color 
         Caption         =   "请选择生成的BAT文件的文字与背景色"
         Height          =   4215
         Left            =   4200
         TabIndex        =   15
         Top             =   120
         Width           =   5895
         Begin VB.Frame Frame_BT_Color_Back 
            Caption         =   "背景色"
            Height          =   1215
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   3855
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   15
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H0000FFFF&
               Height          =   495
               Index           =   14
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00FF00FF&
               Height          =   495
               Index           =   13
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H000000FF&
               Height          =   495
               Index           =   12
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00FFFF00&
               Height          =   495
               Index           =   11
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H0000FF00&
               Height          =   495
               Index           =   10
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00FF0000&
               Height          =   495
               Index           =   9
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00808080&
               Height          =   495
               Index           =   8
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00C0C0C0&
               Height          =   495
               Index           =   7
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00008080&
               Height          =   495
               Index           =   6
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00800080&
               Height          =   495
               Index           =   5
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00000080&
               Height          =   495
               Index           =   4
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00808000&
               Height          =   495
               Index           =   3
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00008000&
               Height          =   495
               Index           =   2
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00800000&
               Height          =   495
               Index           =   1
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Back 
               BackColor       =   &H00000000&
               Height          =   495
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame_BT_Color_Fore 
            Caption         =   "前景色"
            Height          =   1215
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   3855
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   15
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H0000FFFF&
               Height          =   495
               Index           =   14
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00FF00FF&
               Height          =   495
               Index           =   13
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H000000FF&
               Height          =   495
               Index           =   12
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00FFFF00&
               Height          =   495
               Index           =   11
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H0000FF00&
               Height          =   495
               Index           =   10
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00FF0000&
               Height          =   495
               Index           =   9
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00808080&
               Height          =   495
               Index           =   8
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00C0C0C0&
               Height          =   495
               Index           =   7
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00008080&
               Height          =   495
               Index           =   6
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00800080&
               Height          =   495
               Index           =   5
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00000080&
               Height          =   495
               Index           =   4
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00808000&
               Height          =   495
               Index           =   3
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00008000&
               Height          =   495
               Index           =   2
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00800000&
               Height          =   495
               Index           =   1
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton BT_Color_Fore 
               BackColor       =   &H00000000&
               Height          =   495
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.TextBox Text_Show_Color 
            BackColor       =   &H80000012&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   2535
            Left            =   4080
            MultiLine       =   -1  'True
            TabIndex        =   16
            Text            =   "CreatGuide.frx":000C
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame_Theme_Ver 
         Caption         =   "请选择您需要生成的版本"
         Height          =   4215
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   5895
         Begin VB.OptionButton Option_Theme_Ver 
            Caption         =   "Windows通用（注：视觉风格文件不能通用）"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   4500
         End
         Begin VB.OptionButton Option_Theme_Ver 
            Caption         =   "Windows XP / 2003"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1140
            Width           =   4500
         End
         Begin VB.OptionButton Option_Theme_Ver 
            Caption         =   "Windows Vista / 2008"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   4500
         End
         Begin VB.OptionButton Option_Theme_Ver 
            Caption         =   "Windows 7 / 2008 R2"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   11
            Top             =   2460
            Width           =   4500
         End
         Begin VB.OptionButton Option_Theme_Ver 
            Caption         =   "Windows 8"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   3120
            Width           =   4500
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   0
      Left            =   0
      ScaleHeight     =   4380
      ScaleWidth      =   6045
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6045
      Begin VB.Frame Frame_Files 
         Caption         =   "请选择生成何种文件"
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5895
         Begin VB.OptionButton Option_File 
            Caption         =   "Bat文件（用于Win7家庭普通版）"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   2040
            Width           =   4815
         End
         Begin VB.OptionButton Option_File 
            Caption         =   "Windows Theme文件"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   840
            Width           =   4815
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "完成"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "上一步"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "CreatGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i%
Dim B_or_T As Byte '0是theme，1是bat
Dim T_Ver As Byte '系统版本
Dim B_Color_F As Byte '前景色
Dim B_Color_B As Byte '背景色

Dim Tab_Num As Byte
Private Sub change_Tab()
    For i = 0 To picOptions.UBound
        If i = Tab_Num Then
            picOptions(i).Visible = True
        Else
            picOptions(i).Visible = False
        End If
    Next
    If Tab_Num = 0 Then
        '已到达最前的选项,转回到最后的选项
        cmdLast.Enabled = False
        cmdNext.Enabled = True
    ElseIf Tab_Num = 1 Then
        '已到达最后的选项,转回到选项 1
        cmdNext.Enabled = False
        cmdLast.Enabled = True
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
    End If
    If Tab_Num = 1 Then
        If B_or_T = 0 Then
            Frame_Theme_Ver.Visible = True
            Frame_BT_Color.Visible = False
        ElseIf B_or_T = 1 Then
            Frame_BT_Color.Visible = True
            Frame_Theme_Ver.Visible = False
        End If
    End If
End Sub

Private Sub BT_Color_Back_Click(Index As Integer)
    Text_Show_Color.BackColor = BAT_Color(Index)
    B_Color_B = Index
End Sub

Private Sub BT_Color_Fore_Click(Index As Integer)
    Text_Show_Color.ForeColor = BAT_Color(Index)
    B_Color_F = Index
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLast_Click()
    Tab_Num = Tab_Num - 1
    Call change_Tab
End Sub

Private Sub cmdNext_Click()
    Tab_Num = Tab_Num + 1
    Call change_Tab
End Sub

Private Sub cmdOK_Click()
On Error Resume Next '错误返回数字
Dim x%
    Dim save_url As String
    Call WriteIni("Guide", "File", B_or_T, Config_Url)
    Call WriteIni("Guide", "System_Ver", T_Ver, Config_Url)
    Call WriteIni("Guide", "BAT_Color_Fore", B_Color_F, Config_Url)
    Call WriteIni("Guide", "BAT_Color_Back", B_Color_B, Config_Url)
If B_or_T = 0 Then
        CommonDialog1.flags = cdlOFNOverwritePrompt
        CommonDialog1.DialogTitle = Load_Lanuage("保存Windows主题文件", "Public", "CommonDialog_Theme_DialogTitle_Save", Lanuage_Now)
        CommonDialog1.Filter = Load_Lanuage("Windows主题文件", "Public", "CommonDialog_Theme_Filter", Lanuage_Now) & " (*.theme)|*.theme"
        If Main.T_name_E <> "" Then '不为空
            CommonDialog1.filename = Main.T_name_E.text   '打开时默认选择当前文件
        ElseIf Main.T_name_C <> "" Then
            CommonDialog1.filename = Main.T_name_C.text
        Else
            CommonDialog1.filename = Load_Lanuage("我的主题", "Public", "CommonDialog_Theme_filename", Lanuage_Now)
        End If
        CommonDialog1.ShowSave
        If Err.Number = 32755 Then
'            MsgBox "取消"
        Else
            save_url = CommonDialog1.filename
            If T_Ver = 4 Then
                x = MsgBox(Load_Lanuage("到本软件本版本发布为止，还未发现Win8主题被破解，如果您是将Win7主题应用到Win8上，我建议您将视觉风格文件更换为系统默认，以免失去透明效果", "CreatGuide", "Awoke_Theme_Win8", Lanuage_Now), 36, Load_Lanuage("友情提醒", "CreatGuide", "Awoke_Theme_Win8_Tilte", Lanuage_Now))
                If x = 6 Then
                    Call Save_Theme(save_url, T_Ver, True)  '调用theme生成程序
                Else
                    Call Save_Theme(save_url, T_Ver, False) '调用theme生成程序
                End If
            Else
                Call Save_Theme(save_url, T_Ver, False) '调用theme生成程序
            End If
        End If
ElseIf B_or_T = 1 Then
        CommonDialog1.flags = cdlOFNOverwritePrompt
        CommonDialog1.DialogTitle = Load_Lanuage("保存Win7家庭版应用主题的BAT文件", "Public", "CommonDialog_BAT_DialogTitle_Save", Lanuage_Now)
        CommonDialog1.Filter = Load_Lanuage("批处理文件", "Public", "CommonDialog_BAT_Filter", Lanuage_Now) & " (*.bat)|*.bat"
        If Main.T_name_C <> "" Then '不为空
            CommonDialog1.filename = Load_Lanuage("应用", "Public", "CommonDialog_BAT_filename1", Lanuage_Now) & Main.T_name_C.text & Load_Lanuage("到系统", "Public", "CommonDialog_BAT_filename2", Lanuage_Now)
        Else
            CommonDialog1.filename = Load_Lanuage("应用我的主题", "Public", "CommonDialog_BAT_filename3", Lanuage_Now)
        End If
        CommonDialog1.ShowSave
        If Err.Number = 32755 Then
'            MsgBox "取消"
        Else
            save_url = CommonDialog1.filename
            Call Save_Bat(save_url, B_Color_F, B_Color_B) '调用BAT生成程序
        End If
Else
    MsgBox Load_Lanuage("没有您所选择的生成格式", "Public", "No_Files_Filter", Lanuage_Now)
End If
ErrHandler1:
    '用户按“取消”按钮。
End Sub

Private Sub Form_Load()
    Call Change_Lanuage(Lanuage_Now)

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
    For i = 0 To picOptions.UBound
        picOptions(i).BackColor = m_transparencyKey
    Next
    For i = 0 To Option_File.UBound
        Option_File(i).BackColor = m_transparencyKey
    Next
    Frame_Files.BackColor = m_transparencyKey
    Frame_Theme_Ver.BackColor = m_transparencyKey
    Frame_BT_Color.BackColor = m_transparencyKey
    For i = 0 To Option_Theme_Ver.UBound
        Option_Theme_Ver(i).BackColor = m_transparencyKey
    Next
    Frame_BT_Color_Fore.BackColor = m_transparencyKey
    Frame_BT_Color_Back.BackColor = m_transparencyKey
Else
    For i = 0 To picOptions.UBound
        picOptions(i).BackColor = &H8000000F
    Next
    For i = 0 To Option_File.UBound
        Option_File(i).BackColor = &H8000000F
    Next
    Frame_Files.BackColor = &H8000000F
    Frame_Theme_Ver.BackColor = &H8000000F
    Frame_BT_Color.BackColor = &H8000000F
    For i = 0 To Option_Theme_Ver.UBound
        Option_Theme_Ver(i).BackColor = &H8000000F
    Next
    Frame_BT_Color_Fore.BackColor = &H8000000F
    Frame_BT_Color_Back.BackColor = &H8000000F
End If
    Me.Icon = Main.Icon
    '置中窗体
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Tab_Num = 0
If GetFromIni("Guide", "File", Config_Url) <> "" Then
    B_or_T = GetFromIni("Guide", "File", Config_Url)
Else
    B_or_T = 0
End If
If GetFromIni("Guide", "System_Ver", Config_Url) <> "" Then
    T_Ver = GetFromIni("Guide", "System_Ver", Config_Url)
Else
    T_Ver = 0
End If
If GetFromIni("Guide", "BAT_Color_Fore", Config_Url) <> "" Then
    B_Color_F = GetFromIni("Guide", "BAT_Color_Fore", Config_Url)
Else
    B_Color_F = 8
End If
If GetFromIni("Guide", "BAT_Color_Back", Config_Url) <> "" Then
    B_Color_B = GetFromIni("Guide", "BAT_Color_Back", Config_Url)
Else
    B_Color_B = 0
End If
    
    Option_File(B_or_T).value = True
    Option_Theme_Ver(T_Ver).value = True
    BT_Color_Fore(B_Color_F).value = True
    BT_Color_Back(B_Color_B).value = True
    
For i = 0 To picOptions.UBound
    picOptions(i).Top = 0 '移动位置
    picOptions(i).Left = 0
    picOptions(i).Width = 6045
    picOptions(i).Height = 4380
    picOptions(i).Visible = False
Next
Frame_Files.Top = 120
Frame_Files.Left = 120
Frame_Theme_Ver.Top = 120
Frame_Theme_Ver.Left = 120
Frame_BT_Color.Top = 120
Frame_BT_Color.Left = 120
CreatGuide.Width = 6210
CreatGuide.Height = 5340
    Call change_Tab
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

Private Sub Option_File_Click(Index As Integer)
    B_or_T = Index
End Sub

Private Sub Option_Theme_Ver_Click(Index As Integer)
    T_Ver = Index
End Sub
