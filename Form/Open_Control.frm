VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Open_Control 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "手动修改"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_exit 
      Cancel          =   -1  'True
      Caption         =   "退出"
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
      Caption         =   "浏览"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command_scr_open 
      Caption         =   "浏览"
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
      Caption         =   "修改视觉风格"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command_cur_hand 
      Caption         =   "修改鼠标光标"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   680
      Width           =   1935
   End
   Begin VB.CommandButton Command_snd_hand 
      Caption         =   "修改系统音效"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1240
      Width           =   1935
   End
   Begin VB.CommandButton Command_paper_hand 
      Caption         =   "修改壁纸"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command_ico_hand 
      Caption         =   "修改图标"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command_scr_hand 
      Caption         =   "安装屏幕保护程序"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command_glass_hand 
      Caption         =   "修改透明颜色"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command_window_hand 
      Caption         =   "修改窗体颜色和外观"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command_individuation_hand 
      Caption         =   "打开个性化"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label_mss_indro 
      BackStyle       =   0  'Transparent
      Caption         =   "介绍"
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
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0,1") '修改鼠标指针
End Sub

Private Sub Command_exit_Click()
    Unload Me
End Sub

Private Sub Command_glass_Click()
If System_Ver < 6 Then
    MsgBox "检测到您的系统为" & strOSversion & "，您的操作系统版本并无此功能。"
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageColorization") '修改透明颜色
End If
End Sub

Private Sub Command_ico_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
    MsgBox "检测到您的系统为" & strOSversion & "，修改图标请点击自定义桌面按钮。"
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
    MsgBox "检测到您的系统为" & strOSversion & "，老系统我就不知道是什么样子了……"
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
End If
End Sub

Private Sub Command_individuation_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,-1") 'XP打开主题
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
End If
End Sub

Private Sub Command_mss_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = "选择 Windows 视觉样式文件"
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = "Windows 视觉样式文件 (*.msstyles)|*.msstyles"
    If url_mss <> "" Then '不为空
        CommonDialog1.filename = url_to_N(Url_mss_hand) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    Url_mss_hand = CommonDialog1.filename
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_paper_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") 'XP更换壁纸
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageWallpaper") 'Win7更换壁纸
End If
End Sub

Private Sub Command_scr_Click()
Dim x As Integer
If url_scr <> "" And Dir(url_to_N(url_scr)) <> "" Then '不为空且文件存在
    Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '屏幕保护程序
Else
    x = MsgBox("您输入的屏幕保护程序文件地址经检测不存在，是否仍然继续应用到系统？", 4, "文件不存在")
    If x = 6 Then '是
        Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '屏幕保护程序
    End If
End If
End Sub

Private Sub Command_scr_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = "选择屏幕保护程序文件"
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = "屏幕保护程序 (*.scr)|*.scr"
    If url_scr <> "" Then '不为空
        CommonDialog1.filename = url_to_N(Url_scr_hand) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    Url_scr_hand = CommonDialog1.filename
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_snd_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,1") 'XP修改系统音效
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,2") 'Win7修改系统音效
End If
End Sub

Private Sub Command_window_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
    MsgBox "检测到您的系统为" & strOSversion & "，修改窗体颜色请点击高级按钮。"
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
    MsgBox "检测到您的系统为" & strOSversion & "，老系统我就不知道是什么样子了……"
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,advanced,@advanced") '打开窗体设置和外观
End If
End Sub

Private Sub Command_mss_Click()
Dim x As Integer
If url_mss <> "" And Dir(url_to_N(url_mss)) <> "" Then '不为空且文件存在
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
    Call Shell("net stop Themes")
    Call Shell("net start Themes")
Else
    x = MsgBox("您输入的视觉风格文件地址经检测不存在，是否仍然继续应用到系统？", 4, "文件不存在")
    If x = 6 Then '是
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
        Call Shell("net stop Themes")
        Call Shell("net start Themes")
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Icon = Main.Icon
    Label_mss_indro.Caption = "自动应用视觉风格文件可能会重启主题失败，可尝试多点几次。(需管理员权限启动本程序)" & vbCrLf & "如果一直没有应用成功，请检查您是否破解了主题，或者您选择的视觉风格文件的操作系统是否对应"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub
