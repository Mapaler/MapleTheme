VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于我的应用程序"
   ClientHeight    =   3675
   ClientLeft      =   7395
   ClientTop       =   4695
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   2536.55
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdVisitVeb 
      Caption         =   "访问官网"
      Height          =   375
      Left            =   4125
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.Image Image_Icon 
      Height          =   735
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "程序信息"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   210
      TabIndex        =   1
      Top             =   1125
      Width           =   5445
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1410
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "版本"
      Height          =   225
      Left            =   1410
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "警告: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   2
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdVisitVeb_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.mapaler.com/mapletheme", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Form_Paint()
If glass_ok = True Then
'全玻璃↓
    Dim hBrush As Long, m_Rect As rect, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hDC, hBrush)
    GetClientRect Me.hwnd, m_Rect

    FillRect Me.hDC, m_Rect, hBrush
    SelectObject Me.hDC, hBrushOld

    DeleteObject hBrush
'全玻璃↑
End If
End Sub

Private Sub Form_Load()
    Call Change_Lanuage(Lanuage_Now)

'    Me.Caption = "关于 " & App.Title
'    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & App.Revision & " " & Beta
'    lblTitle.Caption = App.Title
'    lblDescription.Caption = "本程序不是自动收集资源到主题文件夹生成主题，而是为已经安装的主题生成家庭版安装BAT" _
'& vbCrLf & "本程序适用于：" & vbCrLf & "主题制作者生成让使用者能在Win7家庭普通版使用主题的BAT" & vbCrLf & "使用主题者自行将没有添加Win7家庭版安装BAT的主题生成安装BAT"
'    lblDisclaimer.Caption = "本程序所有权及使用权归枫谷剑仙所有" & vbCrLf & "程序参考了 樱茶幻萌组@桂叶 出品的"
    Me.Icon = Main.Icon
    Image_Icon.Picture = Main.Icon
If glass_ok = True Then
'全玻璃↓
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
    Exit Sub
ern:
    MsgBox Err.description
'全玻璃↑
End If
End Sub
