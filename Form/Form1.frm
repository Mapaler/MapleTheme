VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "枫谷 Win7 家庭普通版应用主题BAT生成程序"
   ClientHeight    =   11190
   ClientLeft      =   45
   ClientTop       =   -135
   ClientWidth     =   20400
   ForeColor       =   &H80000008&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   20400
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer_Update 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   600
   End
   Begin VB.CommandButton Command_Options 
      Caption         =   "设置"
      Height          =   375
      Left            =   3960
      TabIndex        =   173
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton Command_about 
      Caption         =   "关于"
      Height          =   375
      Left            =   2040
      TabIndex        =   172
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton Command_exit 
      Caption         =   "退出"
      Height          =   375
      Left            =   5880
      TabIndex        =   171
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton Check_ver 
      Caption         =   "检查更新"
      Height          =   375
      Left            =   120
      TabIndex        =   170
      Top             =   6780
      Width           =   1455
   End
   Begin VB.Frame Main_Frame 
      Caption         =   "导出"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   3
      Left            =   8160
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton Command_save_bat 
         Caption         =   "保存BAT(&B)"
         Height          =   375
         Left            =   3840
         TabIndex        =   169
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Change_glass 
         Caption         =   "切换玻璃/普通窗口"
         Height          =   375
         Left            =   120
         TabIndex        =   168
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command_save_theme 
         Caption         =   "保存theme(&T)"
         Height          =   375
         Left            =   2160
         TabIndex        =   167
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Main_Frame 
      Caption         =   "手动应用"
      Height          =   5535
      Index           =   1
      Left            =   11520
      TabIndex        =   5
      Top             =   4560
      Width           =   7455
      Begin VB.CommandButton Command_individuation_hand 
         Caption         =   "打开个性化"
         Height          =   495
         Left            =   5280
         TabIndex        =   163
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command_window_hand 
         Caption         =   "修改窗体颜色和外观"
         Height          =   495
         Left            =   5280
         TabIndex        =   162
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command_glass_hand 
         Caption         =   "修改透明颜色"
         Height          =   495
         Left            =   5280
         TabIndex        =   161
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton Command_scr_hand 
         Caption         =   "安装屏幕保护程序"
         Height          =   495
         Left            =   5280
         TabIndex        =   160
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command_ico_hand 
         Caption         =   "修改图标"
         Height          =   495
         Left            =   1680
         TabIndex        =   159
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command_paper_hand 
         Caption         =   "修改壁纸"
         Height          =   495
         Left            =   1680
         TabIndex        =   158
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command_snd_hand 
         Caption         =   "修改系统音效"
         Height          =   495
         Left            =   1680
         TabIndex        =   157
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command_cur_hand 
         Caption         =   "修改鼠标光标"
         Height          =   495
         Left            =   1680
         TabIndex        =   156
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command_mss_hand 
         Caption         =   "修改视觉风格"
         Height          =   495
         Left            =   5280
         TabIndex        =   155
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Url_mss_hand 
         Height          =   375
         Left            =   120
         TabIndex        =   154
         Top             =   3840
         Width           =   4215
      End
      Begin VB.TextBox Url_scr_hand 
         Height          =   375
         Left            =   120
         TabIndex        =   153
         Top             =   3000
         Width           =   4215
      End
      Begin VB.CommandButton Command_scr_open 
         Caption         =   "浏览"
         Height          =   375
         Left            =   4440
         TabIndex        =   152
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command_mss_open 
         Caption         =   "浏览"
         Height          =   375
         Left            =   4440
         TabIndex        =   151
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label_mss_hand 
         BackStyle       =   0  'Transparent
         Caption         =   "视觉风格文件"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   3600
         Width           =   4215
      End
      Begin VB.Label Label_scr_hand 
         BackStyle       =   0  'Transparent
         Caption         =   "屏幕保护程序文件"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label_mss_indro 
         BackStyle       =   0  'Transparent
         Caption         =   "介绍"
         Height          =   1095
         Left            =   120
         TabIndex        =   164
         Top             =   4320
         Width           =   7215
      End
   End
   Begin VB.Frame Main_Frame 
      Caption         =   "选择主题"
      Height          =   5535
      Index           =   0
      Left            =   13320
      TabIndex        =   4
      Top             =   600
      Width           =   7455
      Begin VB.CommandButton Command_Down_More_Theme 
         Caption         =   "获取更多主题"
         Height          =   615
         Left            =   5280
         TabIndex        =   274
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton Command_Choose_Refresh_Theme 
         Caption         =   "刷新列表"
         Height          =   615
         Left            =   5280
         TabIndex        =   174
         Top             =   3600
         Width           =   2055
      End
      Begin MSComctlLib.ImageList ImageList_Theme 
         Left            =   6840
         Top             =   3360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.CommandButton Command_Choose_Add_Theme 
         Caption         =   "添加列表中没有的主题"
         Height          =   615
         Left            =   5280
         OLEDropMode     =   1  'Manual
         TabIndex        =   148
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command_Choose_Edit_Theme 
         Caption         =   "编辑该主题"
         Height          =   615
         Left            =   5280
         TabIndex        =   147
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton Command_Choose_Aply_Theme 
         Caption         =   "应用到系统"
         Height          =   615
         Left            =   5280
         TabIndex        =   146
         Top             =   720
         Width           =   2055
      End
      Begin MSComctlLib.TreeView TreeView_Theme 
         Height          =   5055
         Left            =   120
         TabIndex        =   145
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   8916
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   7
         ImageList       =   "ImageList_Theme"
         Appearance      =   1
      End
      Begin VB.Label Label_Help_Select_Theme 
         BackStyle       =   0  'Transparent
         Caption         =   "下面列表内是您系统中已经安装的主题，请选择您需要应用的或者编辑的主题"
         Height          =   495
         Left            =   120
         TabIndex        =   149
         Top             =   0
         Width           =   7215
      End
   End
   Begin VB.Frame Frame_Main_Tab 
      BorderStyle     =   0  'None
      Caption         =   "主界面"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton Command_Guide 
         Caption         =   "导出向导"
         Height          =   975
         Left            =   5640
         Picture         =   "Form1.frx":23D2
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   0
         Width           =   1815
      End
      Begin VB.OptionButton Option_Main_Tab 
         Caption         =   "选择主题文件"
         Height          =   975
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":329C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
      Begin VB.OptionButton Option_Main_Tab 
         Caption         =   "手动应用"
         Height          =   975
         Index           =   1
         Left            =   1830
         Picture         =   "Form1.frx":4166
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1815
      End
      Begin VB.OptionButton Option_Main_Tab 
         Caption         =   "编辑主题文件"
         Height          =   975
         Index           =   2
         Left            =   3810
         Picture         =   "Form1.frx":5030
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Main_Frame 
      BackColor       =   &H008080FF&
      Caption         =   "编辑主题"
      Height          =   10095
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   19695
      Begin VB.CommandButton Command_Aply_Now 
         Caption         =   "测试应用效果"
         Height          =   480
         Left            =   0
         TabIndex        =   150
         Top             =   5100
         Width           =   1335
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H00FF8EC7&
         Caption         =   "屏幕保护程序"
         Height          =   5535
         Index           =   6
         Left            =   10920
         TabIndex        =   135
         Top             =   2640
         Width           =   6015
         Begin VB.TextBox scr_wait_min 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3720
            MaxLength       =   5
            TabIndex        =   138
            Text            =   "10"
            Top             =   1080
            Width           =   630
         End
         Begin VB.CommandButton Command_scr 
            Caption         =   "浏览"
            Height          =   375
            Left            =   4680
            TabIndex        =   137
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox url_scr 
            Height          =   375
            Left            =   120
            TabIndex        =   136
            Top             =   480
            Width           =   4455
         End
         Begin MSComCtl2.UpDown UpDown_scr_wait_min 
            Height          =   345
            Left            =   4320
            TabIndex        =   139
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "scr_wait_min"
            BuddyDispid     =   196645
            OrigLeft        =   1680
            OrigTop         =   1950
            OrigRight       =   1935
            OrigBottom      =   2265
            Max             =   9999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label_scr_wait_min 
            BackStyle       =   0  'Transparent
            Caption         =   "分钟"
            Height          =   255
            Left            =   4680
            TabIndex        =   142
            Top             =   1155
            Width           =   495
         End
         Begin VB.Label Label_scr_url 
            BackStyle       =   0  'Transparent
            Caption         =   "屏保文件：（不启用屏保留空即可）"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label_scr_wait 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "等待:"
            Height          =   255
            Left            =   2520
            TabIndex        =   140
            Top             =   1155
            Width           =   1215
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H00FFFFB3&
         Caption         =   "鼠标指针"
         Height          =   5535
         Index           =   4
         Left            =   9480
         TabIndex        =   128
         Top             =   2640
         Width           =   6015
         Begin VB.PictureBox Mouse_box 
            Height          =   3075
            Left            =   240
            ScaleHeight     =   3015
            ScaleWidth      =   5055
            TabIndex        =   240
            Top             =   1200
            Width           =   5120
            Begin VB.VScrollBar VScroll_cur 
               Height          =   3022
               LargeChange     =   200
               Left            =   4800
               Max             =   800
               SmallChange     =   100
               TabIndex        =   241
               Top             =   0
               Value           =   1
               Width           =   255
            End
            Begin VB.Frame Frame_Mouse 
               BorderStyle     =   0  'None
               Caption         =   "鼠标项"
               Height          =   5055
               Left            =   0
               TabIndex        =   242
               Top             =   0
               Width           =   4935
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   14
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   257
                  Top             =   8400
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   13
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   256
                  Top             =   7800
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   12
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   255
                  Top             =   7200
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   11
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   254
                  Top             =   6600
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   10
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   253
                  Top             =   6000
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   9
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   252
                  Top             =   5400
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   8
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   251
                  Top             =   4800
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   7
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   250
                  Top             =   4200
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   6
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   249
                  Top             =   3600
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   5
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   248
                  Top             =   3000
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   4
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   247
                  Top             =   2400
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   3
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   246
                  Top             =   1800
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   2
                  Left            =   4205
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   245
                  Top             =   1200
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   1
                  Left            =   4220
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   244
                  Top             =   600
                  Width           =   510
               End
               Begin VB.PictureBox Pic_Cur 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Height          =   510
                  Index           =   0
                  Left            =   4220
                  ScaleHeight     =   450
                  ScaleWidth      =   450
                  TabIndex        =   243
                  Top             =   20
                  Width           =   510
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "连接选择"
                  Height          =   510
                  Index           =   14
                  Left            =   0
                  TabIndex        =   272
                  Top             =   8400
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "候选"
                  Height          =   510
                  Index           =   13
                  Left            =   0
                  TabIndex        =   271
                  Top             =   7800
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "选定文本"
                  Height          =   510
                  Index           =   5
                  Left            =   0
                  TabIndex        =   270
                  Top             =   3000
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "手写"
                  Height          =   510
                  Index           =   6
                  Left            =   0
                  TabIndex        =   269
                  Top             =   3600
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "不可用"
                  Height          =   510
                  Index           =   7
                  Left            =   0
                  TabIndex        =   268
                  Top             =   4200
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "垂直调整"
                  Height          =   510
                  Index           =   8
                  Left            =   0
                  TabIndex        =   267
                  Top             =   4800
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "水平调整"
                  Height          =   510
                  Index           =   9
                  Left            =   0
                  TabIndex        =   266
                  Top             =   5400
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "对角线1"
                  Height          =   510
                  Index           =   10
                  Left            =   0
                  TabIndex        =   265
                  Top             =   6000
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "移动"
                  Height          =   510
                  Index           =   12
                  Left            =   0
                  TabIndex        =   264
                  Top             =   7200
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "对角线2"
                  Height          =   510
                  Index           =   11
                  Left            =   0
                  TabIndex        =   263
                  Top             =   6600
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "精确定位"
                  Height          =   510
                  Index           =   4
                  Left            =   0
                  TabIndex        =   262
                  Top             =   2400
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "忙"
                  Height          =   510
                  Index           =   3
                  Left            =   0
                  TabIndex        =   261
                  Top             =   1800
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "后台运行"
                  Height          =   510
                  Index           =   2
                  Left            =   0
                  TabIndex        =   260
                  Top             =   1200
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "帮助选择"
                  Height          =   510
                  Index           =   1
                  Left            =   15
                  TabIndex        =   259
                  Top             =   600
                  Width           =   4830
               End
               Begin VB.Label Cur_BG 
                  BackColor       =   &H8000000E&
                  Caption         =   "正常选择"
                  Height          =   510
                  Index           =   0
                  Left            =   0
                  TabIndex        =   258
                  Top             =   0
                  Width           =   4830
               End
            End
         End
         Begin VB.TextBox url_cur 
            Height          =   375
            Left            =   240
            TabIndex        =   133
            Top             =   4320
            Width           =   4095
         End
         Begin VB.Timer Timer_cur 
            Interval        =   100
            Left            =   3720
            Top             =   360
         End
         Begin VB.CommandButton Command_cur 
            Caption         =   "浏览"
            Height          =   375
            Left            =   4440
            TabIndex        =   132
            Top             =   4320
            Width           =   975
         End
         Begin VB.CommandButton cur_default 
            Caption         =   "使用默认值(Win7)"
            Height          =   375
            Left            =   2280
            TabIndex        =   131
            Top             =   4800
            Width           =   2055
         End
         Begin VB.PictureBox ShowCur 
            Height          =   855
            Left            =   4440
            ScaleHeight     =   795
            ScaleWidth      =   795
            TabIndex        =   130
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox name_cur 
            Height          =   375
            Left            =   240
            TabIndex        =   129
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label_cur 
            BackStyle       =   0  'Transparent
            Caption         =   "方案名称"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H0086FF86&
         Caption         =   "图标"
         Height          =   5535
         Index           =   3
         Left            =   8040
         TabIndex        =   103
         Top             =   840
         Width           =   6015
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   5
            Left            =   1200
            TabIndex        =   121
            Top             =   4800
            Width           =   3615
         End
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   4
            Left            =   1200
            TabIndex        =   120
            Top             =   3960
            Width           =   3615
         End
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   119
            Top             =   3120
            Width           =   3615
         End
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   118
            Top             =   2280
            Width           =   3615
         End
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   117
            Top             =   1440
            Width           =   3615
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   5
            Left            =   4920
            TabIndex        =   116
            Top             =   4800
            Width           =   735
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   4
            Left            =   4920
            TabIndex        =   115
            Top             =   3960
            Width           =   735
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   3
            Left            =   4920
            TabIndex        =   114
            Top             =   3120
            Width           =   735
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   2
            Left            =   4920
            TabIndex        =   113
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   112
            Top             =   1440
            Width           =   735
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   5
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   111
            Top             =   4440
            Width           =   800
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   4
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   110
            Top             =   3600
            Width           =   800
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   3
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   109
            Top             =   2760
            Width           =   800
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   2
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   108
            Top             =   1920
            Width           =   800
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   1
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   107
            Top             =   1080
            Width           =   800
         End
         Begin VB.PictureBox Pic_icon 
            AutoRedraw      =   -1  'True
            Height          =   800
            Index           =   0
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   735
            TabIndex        =   106
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton Command_icon 
            Caption         =   "浏览"
            Height          =   375
            Index           =   0
            Left            =   4920
            TabIndex        =   105
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox url_icon 
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   104
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "回收站（满）"
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   127
            Top             =   3720
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "Internet Explorer"
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   126
            ToolTipText     =   "仅对BAT生效"
            Top             =   4560
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "回收站（空）"
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   125
            Top             =   2880
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "网络（网上邻居）"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   124
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "文档（我的文档）"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   123
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label Label_icon 
            BackStyle       =   0  'Transparent
            Caption         =   "计算机（我的电脑）"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   122
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H00FF8E8E&
         Caption         =   "系统音效"
         Height          =   5535
         Index           =   5
         Left            =   7560
         TabIndex        =   91
         Top             =   840
         Width           =   6015
         Begin VB.ComboBox Combo_Sys_Snd 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "Form1.frx":5EFA
            Left            =   120
            List            =   "Form1.frx":5EFC
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton Command_sound 
            Caption         =   "浏览"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4920
            TabIndex        =   98
            Top             =   4440
            Width           =   975
         End
         Begin VB.CommandButton sound_Stop 
            Caption         =   "停止"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3840
            TabIndex        =   97
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton sound_Play 
            Caption         =   "试听"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3840
            TabIndex        =   96
            Top             =   4440
            Width           =   975
         End
         Begin VB.TextBox url_sound 
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   4440
            Width           =   3615
         End
         Begin VB.TextBox sound_name_E 
            Height          =   390
            Left            =   3120
            TabIndex        =   94
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox sound_name_C 
            Height          =   390
            Left            =   120
            TabIndex        =   93
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CheckBox Check_snd 
            Caption         =   "使用已存在的方案"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   5535
         End
         Begin MSComctlLib.TreeView TreeView_Sound 
            Height          =   2655
            Left            =   120
            TabIndex        =   100
            Top             =   1680
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   4683
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "ImageList_sound"
            Appearance      =   1
         End
         Begin MSComctlLib.ImageList ImageList_sound 
            Left            =   5040
            Top             =   4920
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":5EFE
                  Key             =   "sound_a"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":6498
                  Key             =   "sound_b"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":6A32
                  Key             =   "sound_c"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label_sound_name_E 
            BackStyle       =   0  'Transparent
            Caption         =   "声音方案英文简写"
            Height          =   255
            Left            =   3120
            TabIndex        =   102
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Labe_sound_name_C 
            BackStyle       =   0  'Transparent
            Caption         =   "声音方案中文"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H0075FFFF&
         Caption         =   "桌面壁纸"
         Height          =   5535
         Index           =   2
         Left            =   5640
         TabIndex        =   74
         Top             =   1560
         Width           =   6015
         Begin VB.CommandButton Papers_Edit_Clear 
            Caption         =   "全部不选"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4560
            TabIndex        =   83
            Top             =   3840
            Width           =   1095
         End
         Begin VB.CommandButton Papers_Edit_Select_All 
            Caption         =   "选中全部"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4560
            TabIndex        =   82
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox url_paper_files 
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   2280
            Width           =   4335
         End
         Begin VB.CommandButton Command_paper_files 
            Caption         =   "选择"
            Height          =   375
            Left            =   4680
            TabIndex        =   80
            ToolTipText     =   "确定修改当前选择项地址"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox Check_paper_change 
            Caption         =   "无序切换"
            Height          =   255
            Left            =   2520
            TabIndex        =   79
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ComboBox Combo_Paper_Change_Time 
            Height          =   300
            ItemData        =   "Form1.frx":6FCC
            Left            =   2520
            List            =   "Form1.frx":6FCE
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton Command_paper 
            Caption         =   "浏览"
            Height          =   375
            Left            =   4680
            TabIndex        =   77
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox url_paper 
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   480
            Width           =   4335
         End
         Begin VB.PictureBox Picture_paper_TEMP 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1125
            Left            =   240
            ScaleHeight     =   1125
            ScaleMode       =   0  'User
            ScaleWidth      =   1500
            TabIndex        =   75
            Top             =   600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin MSComctlLib.ImageList ImageList_wallpapers 
            Left            =   4320
            Top             =   1080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList ImageList_paper_style 
            Left            =   5040
            Top             =   960
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   64
            ImageHeight     =   41
            MaskColor       =   16711935
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":6FD0
                  Key             =   "填充"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":7BE6
                  Key             =   "适应"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":8540
                  Key             =   "拉伸"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":906E
                  Key             =   "平铺"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":9833
                  Key             =   "居中"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageCombo ImageCombo_paper_style 
            Height          =   705
            Left            =   120
            TabIndex        =   84
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1244
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            OLEDragMode     =   1
            Locked          =   -1  'True
            Text            =   "显示模式"
            ImageList       =   "ImageList_paper_style"
         End
         Begin MSComctlLib.TreeView TreeView_paper 
            Height          =   2655
            Left            =   120
            TabIndex        =   85
            Top             =   2760
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4683
            _Version        =   393217
            HideSelection   =   0   'False
            Style           =   1
            ImageList       =   "ImageList_wallpapers"
            Appearance      =   1
         End
         Begin VB.CheckBox Papers_Edit_Allow 
            Caption         =   "允许编辑图片列表"
            Height          =   615
            Left            =   4560
            TabIndex        =   86
            Top             =   2670
            Width           =   1335
         End
         Begin VB.Label Label_paper_files 
            BackStyle       =   0  'Transparent
            Caption         =   "壁纸幻灯片文件夹："
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   2040
            Width           =   4455
         End
         Begin VB.Label Label_paper_change_time 
            BackStyle       =   0  'Transparent
            Caption         =   "幻灯片切换时间："
            Height          =   255
            Left            =   2520
            TabIndex        =   89
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label_paper_style 
            BackStyle       =   0  'Transparent
            Caption         =   "壁纸显示模式："
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label_paper_index 
            BackStyle       =   0  'Transparent
            Caption         =   "主壁纸文件："
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H0086B7FF&
         Caption         =   "视觉风格"
         Height          =   8655
         Index           =   1
         Left            =   3360
         TabIndex        =   31
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton Command_getcolor 
            Caption         =   "一键取色"
            Height          =   495
            Left            =   3360
            TabIndex        =   32
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Frame System_Color_Tab_Frame 
            BorderStyle     =   0  'None
            Caption         =   "颜色选择"
            Height          =   495
            Left            =   225
            TabIndex        =   33
            Top             =   2280
            Width           =   5565
            Begin VB.OptionButton System_Color_Tab 
               Caption         =   "可视化风格调节"
               Height          =   375
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   120
               Width           =   2775
            End
            Begin VB.OptionButton System_Color_Tab 
               Caption         =   "窗口颜色和外观"
               Height          =   375
               Index           =   1
               Left            =   2790
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   120
               Width           =   2775
            End
         End
         Begin VB.Frame Frame_select_mss 
            Caption         =   "风格选择"
            Height          =   855
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   5535
            Begin VB.OptionButton mss_Classic 
               Caption         =   "Classic"
               Height          =   255
               Left            =   3720
               TabIndex        =   72
               ToolTipText     =   "开发中"
               Top             =   240
               Width           =   1740
            End
            Begin VB.CheckBox Check_Alpha 
               Caption         =   "开启透明"
               Height          =   255
               Left            =   120
               TabIndex        =   71
               ToolTipText     =   "开启Aero透明"
               Top             =   540
               Value           =   1  'Checked
               Width           =   5295
            End
            Begin VB.OptionButton mss_Basic 
               Caption         =   "Basic"
               Height          =   255
               Left            =   1920
               TabIndex        =   70
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton mss_Aero 
               Caption         =   "Aero"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.TextBox url_mss 
            Height          =   375
            Left            =   1440
            TabIndex        =   67
            Top             =   1320
            Width           =   3375
         End
         Begin VB.CommandButton Command_mss 
            Caption         =   "浏览"
            Height          =   375
            Left            =   4920
            TabIndex        =   66
            Top             =   1320
            Width           =   735
         End
         Begin VB.Frame System_Color_Frame 
            Caption         =   "可视化风格调节"
            Height          =   3015
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   2520
            Width           =   5775
            Begin VB.HScrollBar HScroll_ColorizationColorBalance 
               Height          =   255
               LargeChange     =   10
               Left            =   240
               Max             =   100
               TabIndex        =   55
               Top             =   1920
               Width           =   1935
            End
            Begin VB.HScrollBar HScroll_ColorizationAfterglow_alpha 
               Height          =   255
               LargeChange     =   10
               Left            =   3120
               Max             =   255
               TabIndex        =   54
               Top             =   1320
               Width           =   1935
            End
            Begin VB.HScrollBar HScroll_ColorizationColor_alpha 
               Height          =   255
               LargeChange     =   10
               Left            =   3120
               Max             =   255
               TabIndex        =   53
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox Value_ColorizationAfterglow 
               Height          =   270
               Left            =   120
               MaxLength       =   10
               TabIndex        =   52
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox Value_ColorizationColor 
               Height          =   270
               Left            =   120
               MaxLength       =   10
               TabIndex        =   51
               Top             =   600
               Width           =   1095
            End
            Begin VB.HScrollBar HScroll_ColorizationAfterglowBalance 
               Height          =   255
               LargeChange     =   10
               Left            =   2880
               Max             =   100
               TabIndex        =   50
               Top             =   1920
               Width           =   1935
            End
            Begin VB.HScrollBar HScroll_ColorizationBlurBalance 
               Height          =   255
               LargeChange     =   10
               Left            =   240
               Max             =   100
               TabIndex        =   49
               Top             =   2520
               Width           =   1935
            End
            Begin VB.HScrollBar HScroll_ColorizationGlassReflectionIntensity 
               Height          =   255
               LargeChange     =   10
               Left            =   2880
               Max             =   100
               TabIndex        =   48
               Top             =   2520
               Width           =   1935
            End
            Begin VB.TextBox Value_ColorizationColorBalance 
               Height          =   270
               IMEMode         =   3  'DISABLE
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   47
               Top             =   1905
               Width           =   495
            End
            Begin VB.TextBox Value_ColorizationBlurBalance 
               Height          =   270
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   46
               Top             =   2505
               Width           =   495
            End
            Begin VB.TextBox Value_ColorizationAfterglowBalance 
               Height          =   270
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   45
               Top             =   1905
               Width           =   495
            End
            Begin VB.TextBox Value_ColorizationGlassReflectionIntensity 
               Height          =   270
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   44
               Top             =   2520
               Width           =   495
            End
            Begin VB.PictureBox Show_ColorizationColor 
               Height          =   495
               Left            =   2520
               ScaleHeight     =   435
               ScaleWidth      =   435
               TabIndex        =   43
               Top             =   360
               Width           =   495
            End
            Begin VB.PictureBox Show_ColorizationAfterglow 
               Height          =   495
               Left            =   2520
               ScaleHeight     =   435
               ScaleWidth      =   435
               TabIndex        =   42
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label_ColorizationGlassReflectionIntensity 
               BackStyle       =   0  'Transparent
               Caption         =   "Aero条纹数量"
               Height          =   255
               Left            =   2880
               TabIndex        =   65
               ToolTipText     =   "大背景透明度"
               Top             =   2280
               Width           =   2895
            End
            Begin VB.Label Label_ColorizationBlurBalance 
               BackStyle       =   0  'Transparent
               Caption         =   "模糊平衡"
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   2280
               Width           =   2535
            End
            Begin VB.Label Label_ColorizationAfterglowBalance 
               BackStyle       =   0  'Transparent
               Caption         =   "发光颜色平衡"
               Height          =   255
               Left            =   2880
               TabIndex        =   63
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Label_ColorizationColorBalance 
               BackStyle       =   0  'Transparent
               Caption         =   "主颜色平衡"
               Height          =   180
               Left            =   240
               TabIndex        =   62
               Top             =   1680
               Width           =   2535
            End
            Begin VB.Label Label_ColorizationAfterglow_alpha 
               BackStyle       =   0  'Transparent
               Caption         =   "发光颜色透明度"
               Height          =   255
               Left            =   3120
               TabIndex        =   61
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label_ColorizationColor_alpha 
               BackStyle       =   0  'Transparent
               Caption         =   "主颜色透明度"
               Height          =   255
               Left            =   3120
               TabIndex        =   60
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label_ColorizationAfterglow 
               BackStyle       =   0  'Transparent
               Caption         =   "发光颜色"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   1080
               Width           =   2295
            End
            Begin VB.Label Label_ColorizationColor 
               BackStyle       =   0  'Transparent
               Caption         =   "主颜色"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label Value_ColorizationColor_alpha 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               Height          =   255
               Left            =   5160
               TabIndex        =   57
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Value_ColorizationAfterglow_alpha 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               Height          =   255
               Left            =   5160
               TabIndex        =   56
               Top             =   1320
               Width           =   495
            End
         End
         Begin VB.Frame System_Color_Frame 
            Caption         =   "窗口颜色和外观"
            Height          =   3015
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   5520
            Width           =   5775
            Begin VB.PictureBox System_Color_box 
               Height          =   2175
               Left            =   2400
               ScaleHeight     =   2115
               ScaleWidth      =   3315
               TabIndex        =   175
               Top             =   720
               Width           =   3375
               Begin VB.VScrollBar VScroll_System_color 
                  Height          =   2120
                  LargeChange     =   100
                  Left            =   3060
                  Max             =   1000
                  SmallChange     =   50
                  TabIndex        =   176
                  Top             =   0
                  Width           =   255
               End
               Begin VB.Frame Frame_System_Color 
                  BorderStyle     =   0  'None
                  Caption         =   "系统颜色"
                  Height          =   4215
                  Left            =   0
                  TabIndex        =   177
                  Top             =   0
                  Width           =   3135
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   30
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   208
                     Text            =   "255 255 255"
                     Top             =   8950
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   29
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   207
                     Text            =   "255 255 255"
                     Top             =   8710
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   28
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   206
                     Text            =   "255 255 255"
                     Top             =   8365
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   27
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   205
                     Text            =   "255 255 255"
                     Top             =   8110
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   26
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   204
                     Text            =   "255 255 255"
                     Top             =   7765
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   25
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   203
                     Text            =   "255 255 255"
                     Top             =   7525
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   24
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   202
                     Text            =   "255 255 255"
                     Top             =   7180
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   23
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   201
                     Text            =   "255 255 255"
                     Top             =   6895
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   22
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   200
                     Text            =   "255 255 255"
                     Top             =   6550
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   21
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   199
                     Text            =   "255 255 255"
                     Top             =   6310
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   20
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   198
                     Text            =   "255 255 255"
                     Top             =   5965
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   19
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   197
                     Text            =   "255 255 255"
                     Top             =   5710
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   18
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   196
                     Text            =   "255 255 255"
                     Top             =   5365
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   17
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   195
                     Text            =   "255 255 255"
                     Top             =   5125
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   16
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   194
                     Text            =   "255 255 255"
                     Top             =   4780
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   15
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   193
                     Text            =   "255 255 255"
                     Top             =   4495
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   14
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   192
                     Text            =   "255 255 255"
                     Top             =   4150
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   13
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   191
                     Text            =   "255 255 255"
                     Top             =   3910
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   12
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   190
                     Text            =   "255 255 255"
                     Top             =   3565
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   11
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   189
                     Text            =   "255 255 255"
                     Top             =   3310
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   10
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   188
                     Text            =   "255 255 255"
                     Top             =   2965
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   9
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   187
                     Text            =   "255 255 255"
                     Top             =   2725
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   8
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   186
                     Text            =   "255 255 255"
                     Top             =   2380
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   7
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   185
                     Text            =   "255 255 255"
                     Top             =   2110
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   6
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   184
                     Text            =   "255 255 255"
                     Top             =   1765
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   5
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   183
                     Text            =   "255 255 255"
                     Top             =   1525
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   4
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   182
                     Text            =   "255 255 255"
                     Top             =   1180
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   3
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   181
                     Text            =   "255 255 255"
                     Top             =   925
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   2
                     Left            =   1965
                     MaxLength       =   11
                     TabIndex        =   180
                     Text            =   "255 255 255"
                     Top             =   580
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   1
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   179
                     Text            =   "255 255 255"
                     Top             =   340
                     Width           =   1095
                  End
                  Begin VB.TextBox Value_System_Color 
                     Height          =   270
                     Index           =   0
                     Left            =   1970
                     MaxLength       =   11
                     TabIndex        =   178
                     Text            =   "255 255 255"
                     Top             =   0
                     Width           =   1095
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   30
                     Left            =   45
                     TabIndex        =   239
                     Top             =   8965
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   29
                     Left            =   0
                     TabIndex        =   238
                     Top             =   8725
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   28
                     Left            =   45
                     TabIndex        =   237
                     Top             =   8380
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   27
                     Left            =   0
                     TabIndex        =   236
                     Top             =   8125
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   26
                     Left            =   45
                     TabIndex        =   235
                     Top             =   7780
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   25
                     Left            =   0
                     TabIndex        =   234
                     Top             =   7540
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   24
                     Left            =   45
                     TabIndex        =   233
                     Top             =   7195
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   23
                     Left            =   0
                     TabIndex        =   232
                     Top             =   6910
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   22
                     Left            =   45
                     TabIndex        =   231
                     Top             =   6565
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   21
                     Left            =   0
                     TabIndex        =   230
                     Top             =   6325
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   20
                     Left            =   45
                     TabIndex        =   229
                     Top             =   5980
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   19
                     Left            =   0
                     TabIndex        =   228
                     Top             =   5725
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   18
                     Left            =   45
                     TabIndex        =   227
                     Top             =   5380
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   17
                     Left            =   0
                     TabIndex        =   226
                     Top             =   5140
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   16
                     Left            =   45
                     TabIndex        =   225
                     Top             =   4795
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   15
                     Left            =   0
                     TabIndex        =   224
                     Top             =   4510
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   14
                     Left            =   45
                     TabIndex        =   223
                     Top             =   4165
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   13
                     Left            =   0
                     TabIndex        =   222
                     Top             =   3925
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   12
                     Left            =   45
                     TabIndex        =   221
                     Top             =   3580
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   11
                     Left            =   0
                     TabIndex        =   220
                     Top             =   3325
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   10
                     Left            =   45
                     TabIndex        =   219
                     Top             =   2980
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   9
                     Left            =   0
                     TabIndex        =   218
                     Top             =   2740
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   8
                     Left            =   45
                     TabIndex        =   217
                     Top             =   2395
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   7
                     Left            =   0
                     TabIndex        =   216
                     Top             =   2125
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   6
                     Left            =   0
                     TabIndex        =   215
                     Top             =   1780
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   5
                     Left            =   0
                     TabIndex        =   214
                     Top             =   1540
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   4
                     Left            =   45
                     TabIndex        =   213
                     Top             =   1195
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   3
                     Left            =   0
                     TabIndex        =   212
                     Top             =   940
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   2
                     Left            =   45
                     TabIndex        =   211
                     Top             =   595
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   1
                     Left            =   0
                     TabIndex        =   210
                     Top             =   355
                     Width           =   1935
                  End
                  Begin VB.Label Lable_System_Color 
                     BackStyle       =   0  'Transparent
                     Caption         =   "GradientInactiveTitle"
                     Height          =   255
                     Index           =   0
                     Left            =   50
                     TabIndex        =   209
                     Top             =   20
                     Width           =   1935
                  End
               End
            End
            Begin VB.CheckBox Check_insert_system_color 
               Caption         =   "将自定义颜色加入到保存的主题或BAT文件中。（不选则为该风格系统默认值）"
               Height          =   975
               Left            =   120
               TabIndex        =   37
               Top             =   1920
               Width           =   2175
            End
            Begin MSComctlLib.ImageCombo ImageCombo_Classic_Style 
               Height          =   810
               Left            =   120
               TabIndex        =   38
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   1429
               _Version        =   393216
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               OLEDragMode     =   1
               Locked          =   -1  'True
               ImageList       =   "ImageList_Classic_Style"
            End
            Begin MSComctlLib.ImageList ImageList_Classic_Style 
               Left            =   1680
               Top             =   2280
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   48
               ImageHeight     =   48
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   6
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":9FAC
                     Key             =   "p1"
                     Object.Tag             =   "自定义"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":A424
                     Key             =   "p2"
                     Object.Tag             =   "Windows 经典"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":A89C
                     Key             =   "p3"
                     Object.Tag             =   "高对比度 #1"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":AD60
                     Key             =   "p4"
                     Object.Tag             =   "高对比度 #2"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":B264
                     Key             =   "p5"
                     Object.Tag             =   "高对比黑色"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":B6C5
                     Key             =   "p6"
                     Object.Tag             =   "高对比白色"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label_Classic_Style 
               BackStyle       =   0  'Transparent
               Caption         =   "经典风格预设："
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Color_Warn 
               BackStyle       =   0  'Transparent
               Caption         =   "自己编辑颜色可能导致一些奇怪的颜色，请使用一键取色工具"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   5535
            End
         End
         Begin VB.Label Label_mss 
            BackStyle       =   0  'Transparent
            Caption         =   "视觉风格文件"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1395
            Width           =   1335
         End
      End
      Begin VB.Frame Edit_Panel_Frame 
         BackColor       =   &H00BFBFFF&
         Caption         =   "主题信息"
         Height          =   5535
         Index           =   0
         Left            =   1920
         TabIndex        =   16
         Top             =   0
         Width           =   6015
         Begin VB.PictureBox Picture_logo 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   1200
            Left            =   1440
            ScaleHeight     =   1140
            ScaleWidth      =   3540
            TabIndex        =   143
            Top             =   3600
            Width           =   3600
         End
         Begin VB.TextBox Maker_Introduce 
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   1920
            Width           =   5775
         End
         Begin VB.TextBox url_Tlogo 
            Height          =   375
            Left            =   1440
            TabIndex        =   22
            Top             =   3120
            Width           =   3615
         End
         Begin VB.CommandButton Command_Tlogo 
            Caption         =   "选择"
            Height          =   375
            Left            =   5160
            TabIndex        =   21
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox Maker_Web_Url 
            Height          =   375
            Left            =   3000
            TabIndex        =   20
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox Maker_Name 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox T_name_C 
            Height          =   390
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox T_name_E 
            Height          =   375
            Left            =   3000
            TabIndex        =   17
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label_Logo_Preview 
            BackStyle       =   0  'Transparent
            Caption         =   "预览："
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label Label_Maker_Introduce 
            BackStyle       =   0  'Transparent
            Caption         =   "其他版权信息或说明"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   5295
         End
         Begin VB.Label Label_Tlogo 
            BackStyle       =   0  'Transparent
            Caption         =   "主题LOGO："
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label_maker_web 
            BackStyle       =   0  'Transparent
            Caption         =   "网址或个人主页"
            Height          =   255
            Left            =   3120
            TabIndex        =   28
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label_maker 
            BackStyle       =   0  'Transparent
            Caption         =   "主题制作者"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label_TnameC 
            BackStyle       =   0  'Transparent
            Caption         =   "主题中文名称"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label_TnameE 
            BackStyle       =   0  'Transparent
            Caption         =   "主题英文名称"
            Height          =   255
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label_logo_help 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "推荐可以任意透明的PNG格式，LOGO最大显示为240×80像素（在256×256状态下），因此请不要太大"
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   120
            TabIndex        =   24
            Top             =   4920
            Width           =   5775
         End
      End
      Begin VB.Frame Frame_Edit_Panel_Tab 
         BorderStyle     =   0  'None
         Caption         =   "编辑选项"
         Height          =   5415
         Left            =   0
         TabIndex        =   8
         Top             =   -120
         Width           =   1440
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H00FF8EC7&
            Caption         =   "屏幕保护程序"
            Height          =   735
            Index           =   6
            Left            =   0
            Picture         =   "Form1.frx":BA64
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   4440
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H00FF8E8E&
            Caption         =   "系统音效"
            Height          =   735
            Index           =   5
            Left            =   0
            Picture         =   "Form1.frx":C32E
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3720
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H00FFFFB3&
            Caption         =   "鼠标指针"
            Height          =   735
            Index           =   4
            Left            =   0
            Picture         =   "Form1.frx":CBF8
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3000
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H0086FF86&
            Caption         =   "图标"
            Height          =   735
            Index           =   3
            Left            =   0
            Picture         =   "Form1.frx":D4C2
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H0075FFFF&
            Caption         =   "桌面壁纸"
            Height          =   735
            Index           =   2
            Left            =   0
            Picture         =   "Form1.frx":DD8C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H0086B7FF&
            Caption         =   "视觉风格"
            Height          =   735
            Index           =   1
            Left            =   0
            Picture         =   "Form1.frx":E656
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Edit_Panel_Tab 
            BackColor       =   &H00BFBFFF&
            Caption         =   "主题信息"
            Height          =   735
            Index           =   0
            Left            =   0
            Picture         =   "Form1.frx":EF20
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'播放WAV的声明
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_LOOP = &H8 '循环
Const SND_ASYNC = &H1 '异步

Private cAni As New cAniCursor '光标预览
Private cAni2 As New cAniCursor '动画光标
Private cIco As New cAniCursor '图标

Dim ColorizationColor_I As Long, ColorizationAfterglow_I As Long '创建颜色临时储存地址
Dim cur_num As Byte, ico_num As Byte, sound_num As Byte
Dim not_First_Load As Boolean, change_System_color_text As Boolean   '是否是退出

'重画图标列表预览
Public Sub Draw_Ico()
Dim i As Byte
Dim file_path As String
For i = 0 To 5
    If InStr(SysIco(i, 0), ".exe,") <> 0 Then
        file_path = Left(SysIco(i, 0), InStr(SysIco(i, 0), ".exe,") + 3)
    ElseIf InStr(SysIco(i, 0), ".dll,") <> 0 Then
        file_path = Left(SysIco(i, 0), InStr(SysIco(i, 0), ".dll,") + 3)
    Else
        file_path = SysIco(i, 0) '一般
    End If
    
    If file_path <> "" And Dir(url_to_N(file_path)) <> "" Then
        cIco.LoadFromFile url_to_N(SysIco(i, 0))
        cIco.Draw Pic_icon(i).hDC, 0, 0, 48, 48, Pic_icon(i).BackColor
    ElseIf file_path = "" Then
        cIco.LoadFromFile url_to_N(SysIco(i, 3))
        cIco.Draw Pic_icon(i).hDC, 0, 0, 48, 48, Pic_icon(i).BackColor
    Else
        Pic_icon(i).Line (0, 0)-(1000, 1000), Pic_icon(i).BackColor, BF
    End If
    Pic_icon(i).Refresh
Next
End Sub

'重画鼠标列表预览
Public Sub Draw_Cur()
Dim i%
For i = 0 To 14
    If SysCur(i, 0) <> "" And Dir(url_to_N(SysCur(i, 0))) <> "" Then
        cAni.LoadFromFile url_to_N(SysCur(i, 0))
        cAni.Draw Pic_Cur(i).hDC, 2, 2, , , Pic_Cur(i).BackColor
    Else
        Pic_Cur(i).Cls
    End If
    Pic_Cur(i).Refresh
Next
End Sub
'是否用系统默认声音
Private Sub Check_snd_Click()
If Check_snd.value = 0 Then
    Labe_sound_name_C.Enabled = True
    Label_sound_name_E.Enabled = True
    sound_name_C.Enabled = True
    sound_name_E.Enabled = True
    url_sound.Enabled = True
    Command_sound.Enabled = True
    Combo_Sys_Snd.Enabled = False
Else
    Labe_sound_name_C.Enabled = False
    Label_sound_name_E.Enabled = False
    sound_name_C.Enabled = False
    sound_name_E.Enabled = False
    url_sound.Enabled = False
    Command_sound.Enabled = False
    Combo_Sys_Snd.Enabled = True
End If
End Sub

Private Sub Combo_Sys_Snd_Click()
If TreeView_Sound.Nodes.count > 1 Then
Dim i As Integer
    For i = 0 To UBound(Sound, 2)
        Sound(2, i) = Sound(Combo_Sys_Snd.ListIndex + 3, i)
        If Sound(2, i) <> "" Then
            TreeView_Sound.Nodes("s" & i).Image = 2 '将有值的节点小喇叭变白
        Else
            TreeView_Sound.Nodes("s" & i).Image = 0 '将无值的节点小喇叭删掉
        End If
    Next
End If
End Sub

Private Sub Command_Choose_Add_Theme_Click()

Dim DlgInfo As DlgFileInfo
Dim i As Integer
Dim Nodes_num As Integer
Nodes_num = TreeView_Theme.Nodes.count
CommonDialog1.filename = "" '清空文件名
On Error GoTo ErrHandle
'选择文件
With CommonDialog1
    .CancelError = False
    .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
    .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    .DialogTitle = Load_Lanuage("选择Windows主题文件", "Public", "CommonDialog_Theme_DialogTitle_Load", Lanuage_Now)
    .Filter = Load_Lanuage("Windows主题文件", "Public", "CommonDialog_Theme_Filter", Lanuage_Now) & " (*.theme)|*.theme"
    .ShowOpen
    DlgInfo = GetDlgFileInfo(CommonDialog1.filename)
End With
'ReDim Preserve Theme1(0 To DlgInfo.iCount + Nodes_num) '增加数组大小
Dim Paper_Url As String
Dim Root As Node
For i = 1 + Nodes_num To DlgInfo.iCount + Nodes_num
    'Theme1(i) = DlgInfo.sPath & DlgInfo.sFile(i - Nodes_num)
    Theme1.Add DlgInfo.sPath & DlgInfo.sFile(i - Nodes_num)
    Picture_paper_TEMP.Cls
    Paper_Url = url_to_N(GetFromIni("Control Panel\Desktop", "Wallpaper", Theme1(i)))
    Call PaintPng2(Paper_Url, Picture_paper_TEMP.hDC, pWidth, pHeight)
    ImageList_Theme.ListImages.Add , , Picture_paper_TEMP.Image
    Set Root = TreeView_Theme.Nodes.Add(, , "t" & i, DlgInfo.sFile(i - Nodes_num), i)
Next
ErrHandle:
' 按了“取消”按钮
End Sub
Private Function GetFileName_L(ByVal FileURL As String)
    Dim fname As String
    fname = Mid(FileURL, InStrRev(FileURL, "\") + 1, InStrRev(FileURL, ".") - InStrRev(FileURL, "\") - 1)
    GetFileName_L = fname
End Function
Private Sub Command_Choose_Add_Theme_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim Nodes_num As Integer
    Nodes_num = TreeView_Theme.Nodes.count
    'ReDim Preserve Theme1(0 To data.Files.count + Nodes_num) '增加数组大小
    Dim Paper_Url As String
    Dim Root As Node
    
    For i = 1 + Nodes_num To data.Files.count + Nodes_num
        'Theme1(i) = data.Files(i - Nodes_num)
        Theme1.Add data.Files(i - Nodes_num)
        Picture_paper_TEMP.Cls
        Paper_Url = url_to_N(GetFromIni("Control Panel\Desktop", "Wallpaper", Theme1(i)))
        Call PaintPng2(Paper_Url, Picture_paper_TEMP.hDC, pWidth, pHeight)
        ImageList_Theme.ListImages.Add , , Picture_paper_TEMP.Image
        Set Root = TreeView_Theme.Nodes.Add(, , "t" & i, GetFileName_L(data.Files(i - Nodes_num)), i)
    Next
End Sub

Private Sub Command_Choose_Aply_Theme_Click()
Dim i%, j%, x%
Dim Select_Num As Integer
If TreeView_Theme.Nodes.count <> 0 Then '检查是否有主题
    TreeView_Theme.SetFocus
    If InStr(TreeView_Theme.SelectedItem.Key, "t") = 1 Then
        Select_Num = Mid(TreeView_Theme.SelectedItem.Key, 2)
        Call Load_theme(Theme1(Select_Num))
          '调用读取主题
        Cur_BG_Click (0)
        Draw_Ico
        Draw_Cur
            Call Shell("net stop Themes")
    
        Call Aply_Theme(Theme1(Select_Num))   '调用应用主题子程序
    
            Call Shell("RunDll32.exe user32.DLL, UpdatePerUserSystemParameters")
            Call Shell("net start Themes")
            sndPlaySound url_to_N(Get_dll_text(GetString(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\.Default\ChangeTheme\.Current", vbNullString))), SND_ASYNC '播放一次，SND_ASYNC Or SND_LOOP '循环播放
            MsgBox Load_Lanuage("修改成功，部分设置需要注销重新登入后显示", "Main", "Use_Theme_OK", Lanuage_Now)
            
    'If url_mss <> "" And Dir(url_to_N(url_mss)) <> "" Then '不为空且文件存在
    '    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
    '    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
    '    Call Shell("net stop Themes")
    '    Call Shell("net start Themes")
    'Else
    '    x = MsgBox("您输入的视觉风格文件地址经检测不存在，是否仍然继续应用到系统？", 4, "文件不存在")
    '    If x = 6 Then '是
    '        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
    '        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
    '        Call Shell("net stop Themes")
    '        Call Shell("net start Themes")
    '    End If
    'End If
    End If

End If '检查是否有主题
End Sub

Private Sub Command_Choose_Edit_Theme_Click()
Dim i As Integer
Dim Select_Num As Integer
If TreeView_Theme.Nodes.count <> 0 Then
    TreeView_Theme.SetFocus
    If InStr(TreeView_Theme.SelectedItem.Key, "t") = 1 Then
        Select_Num = Mid(TreeView_Theme.SelectedItem.Key, 2)
        Call Load_theme(Theme1(Select_Num))
          '调用读取主题
        Cur_BG_Click (0)
        Draw_Ico
        Draw_Cur
        Option_Main_Tab(2).value = True
    End If
End If
End Sub

Private Sub Command_Choose_Refresh_Theme_Click()
Call Refresh_Theme
End Sub

Private Sub Command_cur_hand_Click()
'修改鼠标指针
If System_Ver < 6.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0,1")
Else
    Call Shell("control.exe /name Microsoft.Mouse /page 1")
End If
End Sub

Private Sub Command_Down_More_Theme_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.comicdd.com/?fromuid=301742", vbNullString, vbNullString, SW_SHOWNORMAL
'    ShellExecute Me.hWnd, vbNullString, "http://www.mapaler.com/mapletheme/moretheme.php?lan=zh-CN", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Command_glass_hand_Click()
If System_Ver < 6 Then
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，您的操作系统版本并无此功能。", "Main", "My_System2", Lanuage_Now)
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageColorization") '修改透明颜色
End If
End Sub

Private Sub Command_ico_hand_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，修改图标请点击自定义桌面按钮。", "Main", "My_System3", Lanuage_Now)
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，老系统我就不知道是什么样子了……反正我也不支持这个……", "Main", "My_System4", Lanuage_Now)
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") '修改图标
End If
End Sub

Private Sub Command_individuation_hand_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,-1") 'XP打开主题
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
End If
End Sub

Private Sub Command_Aply_Now_Click()
Dim save_url As String
save_url = Environ("temp") & "\MapleTheme_Temp.theme"
Call Save_Theme(save_url, 0, False) '调用theme生成程序
Call Shell("net stop Themes")
Call Aply_Theme(save_url)   '调用应用主题子程序
Call Shell("RunDll32.exe user32.DLL, UpdatePerUserSystemParameters")
Call Shell("net start Themes")
End Sub

'选择视觉风格文件
Private Sub Command_mss_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择 Windows 视觉样式文件", "Public", "CommonDialog_Mss_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("Windows 视觉样式文件", "Public", "CommonDialog_Mss_Filter", Lanuage_Now) & " (*.msstyles)|*.msstyles"
    If url_mss <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_mss) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_mss = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_mss_hand_Click()
Dim x As Integer
If url_mss <> "" And Dir(url_to_N(url_mss)) <> "" Then '不为空且文件存在
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
    Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
    Call Shell("net stop Themes")
    Call Shell("net start Themes")
Else
    x = MsgBox(Load_Lanuage("您输入的视觉风格文件地址经检测不存在，是否仍然继续应用到系统？", "Main", "Load_Mss_Fail", Lanuage_Now), 4, Load_Lanuage("文件不存在", "Main", "Load_Mss_Fail_Title", Lanuage_Now))
    If x = 6 Then '是
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", url_to_S(Url_mss_hand))
        Call SetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1")
        Call Shell("net stop Themes")
        Call Shell("net start Themes")
    End If
End If
End Sub

Private Sub Command_mss_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择 Windows 视觉样式文件", "Public", "CommonDialog_Mss_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = Load_Lanuage("Windows 视觉样式文件", "Public", "CommonDialog_Mss_Filter", Lanuage_Now) & " (*.msstyles)|*.msstyles"
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

Private Sub Command_Options_Click()
    Options.Show 1
End Sub

'选择壁纸文件
Private Sub Command_paper_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择壁纸图片", "Public", "CommonDialog_Paper_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("图片文件", "Public", "CommonDialog_Paper_Filter", Lanuage_Now) & "|*.png;*.jpg;*.jpeg;*.bmp;*.gif"
    If url_paper <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_paper) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_paper = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_paper_files_Click()
Dim url_temp As String
    url_temp = url_to_N(url_paper_files)
    url_temp = url_to_P(BrowseForFolderByPath(url_temp, Load_Lanuage("请选择壁纸文件夹", "Public", "BrowseForFolder_Paper", Lanuage_Now), Me))
    If url_temp <> "" Then
        url_paper_files = url_temp
    End If
End Sub

Private Sub Command_paper_hand_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0") 'XP更换壁纸
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageWallpaper") 'Win7更换壁纸
End If
End Sub

'保存BAT
Private Sub Command_save_bat_Click()
Dim save_url As String
'On Error GoTo ErrHandler
    CommonDialog1.flags = cdlOFNOverwritePrompt
    CommonDialog1.DialogTitle = "保存Win7家庭版应用主题的BAT文件"
    CommonDialog1.Filter = "批处理文件 (*.bat)|*.bat"
    If T_name_C <> "" Then '不为空
        CommonDialog1.filename = "应用 " & T_name_C.text & " 到系统"  '打开时默认选择当前文件
    Else
        CommonDialog1.filename = "应用我的主题"
    End If
    CommonDialog1.ShowSave
    save_url = CommonDialog1.filename
    Save_Bat (save_url) '调用BAT生成程序
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_save_theme_Click()
Dim save_url As String
On Error GoTo ErrHandler
    CommonDialog1.flags = cdlOFNOverwritePrompt
    CommonDialog1.DialogTitle = "保存Win7主题文件"
    CommonDialog1.Filter = "Win7主题文件 (*.theme)|*.theme"
    If T_name_E <> "" Then '不为空
        CommonDialog1.filename = T_name_E.text   '打开时默认选择当前文件
    ElseIf T_name_C <> "" Then
        CommonDialog1.filename = T_name_C.text
    Else
        CommonDialog1.filename = "我的主题"
    End If
    CommonDialog1.ShowSave
    save_url = CommonDialog1.filename
    Save_Theme (save_url) '调用BAT生成程序
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

'选择屏保文件
Private Sub Command_scr_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择屏幕保护程序文件", "Public", "CommonDialog_Scr_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("屏幕保护程序文件", "Public", "CommonDialog_Scr_Filter", Lanuage_Now) & " (*.scr)|*.scr"
    If url_scr <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_scr) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_scr = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_scr_hand_Click()
Dim x As Integer
If url_scr <> "" And Dir(url_to_N(url_scr)) <> "" Then '不为空且文件存在
    Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '屏幕保护程序
Else
    Call Shell("rundll32.exe desk.cpl,InstallScreenSaver " & Url_scr_hand) '屏幕保护程序
End If
End Sub

Private Sub Command_scr_open_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择屏幕保护程序文件", "Public", "CommonDialog_Scr_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.InitDir = "%SystemRoot%\Resources\Themes"
    CommonDialog1.Filter = Load_Lanuage("屏幕保护程序文件", "Public", "CommonDialog_Scr_Filter", Lanuage_Now) & " (*.scr)|*.scr"
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

Private Sub Command_snd_hand_Click()
If System_Ver < 6 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,1") 'XP修改系统音效
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,2") 'Win7修改系统音效
End If
End Sub

'选择声音文件
Private Sub Command_sound_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("浏览新的", "Public", "CommonDialog_Snd_DialogTitle_Load1", Lanuage_Now) & " " & Sound(1, sound_num) & " " & Load_Lanuage("声音", "Public", "CommonDialog_Snd_DialogTitle_Load2", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("波形文件", "Public", "CommonDialog_Snd_Filter", Lanuage_Now) & " (*.wav)|*.wav"
    If url_sound <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_sound) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_sound = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

'选择主题logo文件
Private Sub Command_Tlogo_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择主题LOGO图片", "Public", "CommonDialog_Logo_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("图片文件", "Public", "CommonDialog_Paper_Filter", Lanuage_Now) & "|*.png;*.jpg;*.jpeg;*.bmp;*.gif"
    If url_Tlogo <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_Tlogo) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_Tlogo = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Command_window_hand_Click()
If System_Ver < 6 And System_Ver >= 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，修改窗体颜色请点击高级按钮。", "Main", "My_System5", Lanuage_Now)
ElseIf System_Ver < 5.1 Then
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2") '打开个性化
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，老系统我就不知道是什么样子了……反正我也不支持这个……", "Main", "My_System4", Lanuage_Now)
Else
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,advanced,@advanced") '打开窗体设置和外观
End If
End Sub

Private Sub Command_Guide_Click()
If CreatGuide.Visible = False Then
    CreatGuide.Show 1 '加了1使其他的不能操作
Else
    Unload CreatGuide
End If
End Sub

Private Sub Edit_Panel_Tab_Click(Index As Integer)
Dim i As Integer
For i = 0 To Edit_Panel_Tab.UBound
    If i <> Index Then
        Edit_Panel_Frame(i).Visible = False
        Edit_Panel_Tab(i).Width = 81 * 15
    Else
        Edit_Panel_Frame(i).Visible = True
        Edit_Panel_Tab(i).Width = (81 + 8) * 15 '加宽
    End If
Next
End Sub

'改变风格预设
Private Sub ImageCombo_Classic_Style_Change()
'当不是在文本框按键修改时更换值不存入
If change_System_color_text = False Then
    Dim i As Byte
    Dim item As String
    item = ImageCombo_Classic_Style.SelectedItem.Index
    If item = 1 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 1)
        Next i
    ElseIf item = 2 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 2)
        Next i
    ElseIf item = 3 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 3)
        Next i
    ElseIf item = 4 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 4)
        Next i
    ElseIf item = 5 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 5)
        Next i
    ElseIf item = 6 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 6)
        Next i
    Else
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 2)
        Next i
    End If
End If
If item = 1 Then
        Check_insert_system_color.value = 1
        Check_insert_system_color.Enabled = False
Else
        Check_insert_system_color.Enabled = True
End If
End Sub
Private Sub ImageCombo_Classic_Style_Click()
'当不是在文本框按键修改时更换值不存入
If change_System_color_text = False Then
    Dim i As Byte
    Dim item As String
    item = ImageCombo_Classic_Style.SelectedItem.Index
    If item = 1 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 1)
        Next i
    Get_color.Show
    ElseIf item = 2 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 2)
        Next i
    ElseIf item = 3 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 3)
        Next i
    ElseIf item = 4 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 4)
        Next i
    ElseIf item = 5 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 5)
        Next i
    ElseIf item = 6 Then
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 6)
        Next i
    Else
        For i = 0 To 30
            Value_System_Color(i).text = SysColors(i, 2)
        Next i
    End If
End If
If item = 1 Then
        Check_insert_system_color.value = 1
        Check_insert_system_color.Enabled = False
Else
        Check_insert_system_color.Enabled = True
End If
End Sub



Private Sub Option_Main_Tab_Click(Index As Integer)
Dim i As Integer
For i = 0 To Main_Frame.UBound
    If i <> Index Then
        Main_Frame(i).Visible = False
    Else
        Main_Frame(i).Visible = True
    End If
Next
Call WriteIni("Option", "Main_Tab", Index, Config_Url)
End Sub
'壁纸编辑，是否允许
Private Sub Papers_Edit_Allow_Click()
Dim i As Integer
    If Papers_Edit_Allow.value = 0 Then
    
        Papers_Edit_Select_All.Enabled = False
        Papers_Edit_Clear.Enabled = False
        
        TreeView_paper.Checkboxes = False
    Else
        TreeView_paper.Checkboxes = True
        For i = 1 To TreeView_paper.Nodes.count
            TreeView_paper.Nodes(i).Checked = True '壁纸全选中
        Next
        Papers_Edit_Select_All.Enabled = True
        Papers_Edit_Clear.Enabled = True
    End If
End Sub
'壁纸编辑，清空
Private Sub Papers_Edit_Clear_Click()
Dim i%
    For i = 1 To TreeView_paper.Nodes.count
        TreeView_paper.Nodes(i).Checked = False '壁纸全选中
    Next
End Sub

Private Sub Papers_Edit_Select_All_Click()
Dim i%
    For i = 1 To TreeView_paper.Nodes.count
        TreeView_paper.Nodes(i).Checked = True '壁纸全选中
    Next
End Sub

'屏幕保护程序等待时间
Private Sub scr_wait_min_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub scr_wait_min_Change()
Dim text_temp As Variant
text_temp = text_to_num(scr_wait_min)
If text_temp = "" Then
    text_temp = 0
ElseIf text_temp >= 0 And text_temp <= 9999 Then
    text_temp = text_temp
ElseIf text_temp < 0 Then
    text_temp = 0
ElseIf text_temp > 100 Then
    text_temp = 9999
End If
scr_wait_min = text_temp
End Sub



Private Sub System_Color_Tab_Click(Index As Integer)
Dim i As Integer
For i = 0 To System_Color_Frame.UBound
    If i <> Index Then
        System_Color_Frame(i).Visible = False
'        Edit_Panel_Tab(i).Width = 81 * 15
    Else
        System_Color_Frame(i).Visible = True
'        Edit_Panel_Tab(i).Width = (81 + 8) * 15 '加宽
    End If
Next
End Sub

Private Sub Timer_Update_Timer()
If Timer_Update.Enabled = True Then
    Call CheckVer(UpdataURL, Auto_Update, Me)
End If
Auto_Update = False
Timer_Update.Enabled = False
End Sub

Private Sub url_paper_files_Change()
If Papers_Edit_Allow.value = 0 Or TreeView_paper.Nodes.count = 0 Then '当编辑未选中或者是没有东西的时候
    TreeView_paper.Nodes.Clear '清除以前的节点
    ImageList_wallpapers.ListImages.Clear '清除以前的图片
    'Erase PaperFileName '先清空数组
    Set PaperFileName = New Collection  '先清空数组
    
    Dim i As Integer
    Call GetFileName(url_to_N(url_paper_files), "bmp,jpg,jpeg,gif,png", PaperFileName) '获取文件列表
'    On Error GoTo ErrHandler
        '将所有图片载入imagelist
        New_List = "" '先清空New_List
        'For i = 1 To UBound(PaperFileName)
        For i = 1 To PaperFileName.count
            Picture_paper_TEMP.Cls
            Call PaintPng2(PaperFileName(i), Picture_paper_TEMP.hDC, pWidth, pHeight)
            ImageList_wallpapers.ListImages.Add , , Picture_paper_TEMP.Image
            
            'If i < UBound(PaperFileName) Then
            If i < PaperFileName.count Then
                New_List = New_List & PaperFileName(i) & vbCrLf
            Else
                New_List = New_List & PaperFileName(i)
            End If
        Next
        


        '添加listview节点
        Dim Root As Node
        With TreeView_paper.Nodes
            'For i = 1 To UBound(PaperFileName)
            For i = 1 To PaperFileName.count
                Set Root = .Add(, , "p" & i, Mid$(PaperFileName(i), InStrRev(PaperFileName(i), "\") + 1), i)
            Next
        End With
'ErrHandler:
    '下标错误（没有图片）
    Exit Sub
    
    '默认选中第一个
    'TreeView_paper.Nodes(1).Selected = True
End If
End Sub

Private Sub url_Tlogo_Change()
'加载主题logo预览
If url_Tlogo <> "" And Dir(url_to_N(url_Tlogo)) <> "" Then '不为空且文件存在
    Picture_logo.Cls
    Call PaintPng2(url_to_N(url_Tlogo), Picture_logo.hDC, 240, 80)
Else
    Picture_logo.Cls
End If
End Sub

'选择主颜色
Private Sub Show_ColorizationColor_Click()
Dim i%
Rem 读取颜色
On Error GoTo ErrHandler
    CommonDialog1.Color = ColorizationColor_I
    CommonDialog1.ShowColor
    ColorizationColor_I = RGB_To_BGR(CommonDialog1.Color)
    Value_ColorizationColor = "0x" & x10_to_x16(HScroll_ColorizationColor_alpha.value, 2) & x10_to_x16(ColorizationColor_I, 6)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Value_ColorizationColor_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Or KeyAscii = 88 Or KeyAscii = 120 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationColor_Change()
Dim Color_Alpha As Byte
Value_ColorizationColor = text_to_color(Value_ColorizationColor)
If Color_ok(Value_ColorizationColor) = 10 Then
    Color_Alpha = x16_to_x10(Mid$(Value_ColorizationColor, 3, 2))
    ColorizationColor_I = x16_to_x10(Mid$(Value_ColorizationColor, 5, 6))
    HScroll_ColorizationColor_alpha.value = x16_to_x10(Mid$(Value_ColorizationColor, 3, 2))
    Value_ColorizationColor_alpha.Caption = Round(x16_to_x10(Mid$(Value_ColorizationColor, 3, 2)) / 255 * 100, 1) & "%"
ElseIf Color_ok(Value_ColorizationColor) = 8 Then
    Color_Alpha = x16_to_x10(Mid$(Value_ColorizationColor, 1, 2))
    ColorizationColor_I = x16_to_x10(Mid$(Value_ColorizationColor, 3, 6))
    HScroll_ColorizationColor_alpha.value = x16_to_x10(Mid$(Value_ColorizationColor, 1, 2))
    Value_ColorizationColor_alpha.Caption = Round(x16_to_x10(Mid$(Value_ColorizationColor, 1, 2)) / 255 * 100, 1) & "%"
End If
Show_ColorizationColor.BackColor = RGB_To_BGR_Alpha(ColorizationColor_I, Aplha_Back_Color, Color_Alpha)
End Sub
'主颜色透明度
Private Sub HScroll_ColorizationColor_alpha_Change()
    Value_ColorizationColor = "0x" & x10_to_x16(HScroll_ColorizationColor_alpha.value, 2) & x10_to_x16(ColorizationColor_I, 6)
End Sub
Private Sub HScroll_ColorizationColor_alpha_Scroll()
    Value_ColorizationColor = "0x" & x10_to_x16(HScroll_ColorizationColor_alpha.value, 2) & x10_to_x16(ColorizationColor_I, 6)
End Sub
'选择发光颜色
Private Sub Show_ColorizationAfterglow_Click()
Dim i%
Rem 读取颜色
On Error GoTo ErrHandler
    CommonDialog1.Color = ColorizationAfterglow_I
    CommonDialog1.ShowColor
    ColorizationAfterglow_I = RGB_To_BGR(CommonDialog1.Color)
    Value_ColorizationAfterglow = "0x" & x10_to_x16(HScroll_ColorizationAfterglow_alpha.value, 2) & x10_to_x16(ColorizationAfterglow_I, 6)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Value_ColorizationAfterglow_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Or KeyAscii = 88 Or KeyAscii = 120 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationAfterglow_Change()
Dim Color_Alpha As Byte
Value_ColorizationAfterglow = text_to_color(Value_ColorizationAfterglow)
If Color_ok(Value_ColorizationAfterglow) = 10 Then
    Color_Alpha = x16_to_x10(Mid$(Value_ColorizationAfterglow, 3, 2))
    ColorizationAfterglow_I = x16_to_x10(Mid$(Value_ColorizationAfterglow, 5, 6))
    HScroll_ColorizationAfterglow_alpha.value = x16_to_x10(Mid$(Value_ColorizationAfterglow, 3, 2))
    Value_ColorizationAfterglow_alpha.Caption = Round(x16_to_x10(Mid$(Value_ColorizationAfterglow, 3, 2)) / 255 * 100, 1) & "%"
ElseIf Color_ok(Value_ColorizationAfterglow) = 8 Then
    Color_Alpha = x16_to_x10(Mid$(Value_ColorizationAfterglow, 1, 2))
    ColorizationAfterglow_I = x16_to_x10(Mid$(Value_ColorizationAfterglow, 3, 6))
    HScroll_ColorizationAfterglow_alpha.value = x16_to_x10(Mid$(Value_ColorizationAfterglow, 1, 2))
    Value_ColorizationAfterglow_alpha.Caption = Round(x16_to_x10(Mid$(Value_ColorizationAfterglow, 1, 2)) / 255 * 100, 1) & "%"
End If
Show_ColorizationAfterglow.BackColor = RGB_To_BGR_Alpha(ColorizationAfterglow_I, Aplha_Back_Color, Color_Alpha)
End Sub
'发光颜色透明度
Private Sub HScroll_ColorizationAfterglow_alpha_Change()
    Value_ColorizationAfterglow = "0x" & x10_to_x16(HScroll_ColorizationAfterglow_alpha.value, 2) & x10_to_x16(ColorizationAfterglow_I, 6)
End Sub
Private Sub HScroll_ColorizationAfterglow_alpha_Scroll()
    Value_ColorizationAfterglow = "0x" & x10_to_x16(HScroll_ColorizationAfterglow_alpha.value, 2) & x10_to_x16(ColorizationAfterglow_I, 6)
End Sub


'主颜色平衡
Private Sub HScroll_ColorizationColorBalance_Change()
    Value_ColorizationColorBalance = HScroll_ColorizationColorBalance.value
End Sub
Private Sub HScroll_ColorizationColorBalance_Scroll()
    Value_ColorizationColorBalance = HScroll_ColorizationColorBalance.value
End Sub
Private Sub Value_ColorizationColorBalance_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationColorBalance_Change()
Dim text_temp As Variant
text_temp = text_to_num(Value_ColorizationColorBalance)
If text_temp = "" Then
    HScroll_ColorizationColorBalance.value = 0
ElseIf text_temp >= 0 And text_temp <= 100 Then
    HScroll_ColorizationColorBalance.value = text_temp
ElseIf text_temp < 0 Then
    HScroll_ColorizationColorBalance.value = 0
ElseIf text_temp > 100 Then
    HScroll_ColorizationColorBalance.value = 100
End If
Value_ColorizationColorBalance = text_temp
End Sub

'发光颜色平衡
Private Sub HScroll_ColorizationAfterglowBalance_Change()
    Value_ColorizationAfterglowBalance = HScroll_ColorizationAfterglowBalance.value
End Sub
Private Sub HScroll_ColorizationAfterglowBalance_Scroll()
    Value_ColorizationAfterglowBalance = HScroll_ColorizationAfterglowBalance.value
End Sub
Private Sub Value_ColorizationAfterglowBalance_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationAfterglowBalance_Change()
Dim text_temp As Variant
text_temp = text_to_num(Value_ColorizationAfterglowBalance)
If text_temp = "" Then
    HScroll_ColorizationAfterglowBalance.value = 0
ElseIf text_temp >= 0 And text_temp <= 100 Then
    HScroll_ColorizationAfterglowBalance.value = text_temp
ElseIf text_temp < 0 Then
    HScroll_ColorizationAfterglowBalance.value = 0
ElseIf text_temp > 100 Then
    HScroll_ColorizationAfterglowBalance.value = 100
End If
Value_ColorizationAfterglowBalance = text_temp
End Sub
'模糊平衡
Private Sub HScroll_ColorizationBlurBalance_Change()
    Value_ColorizationBlurBalance = HScroll_ColorizationBlurBalance.value
End Sub
Private Sub HScroll_ColorizationBlurBalance_Scroll()
    Value_ColorizationBlurBalance = HScroll_ColorizationBlurBalance.value
End Sub
Private Sub Value_ColorizationBlurBalance_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationBlurBalance_Change()
Dim text_temp As Variant
text_temp = text_to_num(Value_ColorizationBlurBalance)
If text_temp = "" Then
    HScroll_ColorizationBlurBalance.value = 0
ElseIf text_temp >= 0 And text_temp <= 100 Then
    HScroll_ColorizationBlurBalance.value = text_temp
ElseIf text_temp < 0 Then
    HScroll_ColorizationBlurBalance.value = 0
ElseIf text_temp > 100 Then
    HScroll_ColorizationBlurBalance.value = 100
End If
Value_ColorizationBlurBalance = text_temp
End Sub
'Aero条纹数量
Private Sub HScroll_ColorizationGlassReflectionIntensity_Change()
    Value_ColorizationGlassReflectionIntensity = HScroll_ColorizationGlassReflectionIntensity.value
End Sub
Private Sub HScroll_ColorizationGlassReflectionIntensity_Scroll()
    Value_ColorizationGlassReflectionIntensity = HScroll_ColorizationGlassReflectionIntensity.value
End Sub
Private Sub Value_ColorizationGlassReflectionIntensity_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub Value_ColorizationGlassReflectionIntensity_Change()
Dim text_temp As Variant
text_temp = text_to_num(Value_ColorizationGlassReflectionIntensity)
If text_temp = "" Then
    HScroll_ColorizationGlassReflectionIntensity.value = 0
ElseIf text_temp >= 0 And text_temp <= 100 Then
    HScroll_ColorizationGlassReflectionIntensity.value = text_temp
ElseIf text_temp < 0 Then
    HScroll_ColorizationGlassReflectionIntensity.value = 0
ElseIf text_temp > 100 Then
    HScroll_ColorizationGlassReflectionIntensity.value = 100
End If
Value_ColorizationGlassReflectionIntensity = text_temp
End Sub

'更改壁纸显示模式
Private Sub ImageCombo_paper_style_Change()
Dim item As ComboItem
Set item = ImageCombo_paper_style.SelectedItem
If item.Key = "填充" Then
    WallpaperStyle_value = 10
    TileWallpaper_value = 0
ElseIf item.Key = "适应" Then
    WallpaperStyle_value = 6
    TileWallpaper_value = 0
ElseIf item.Key = "拉伸" Then
    WallpaperStyle_value = 2
    TileWallpaper_value = 0
ElseIf item.Key = "平铺" Then
    WallpaperStyle_value = 0
    TileWallpaper_value = 1
ElseIf item.Key = "居中" Then
    WallpaperStyle_value = 0
    TileWallpaper_value = 0
End If
End Sub
Private Sub ImageCombo_paper_style_Click()
Dim item As ComboItem
Set item = ImageCombo_paper_style.SelectedItem
If item.Key = "填充" Then
    WallpaperStyle_value = 10
    TileWallpaper_value = 0
ElseIf item.Key = "适应" Then
    WallpaperStyle_value = 6
    TileWallpaper_value = 0
ElseIf item.Key = "拉伸" Then
    WallpaperStyle_value = 2
    TileWallpaper_value = 0
ElseIf item.Key = "平铺" Then
    WallpaperStyle_value = 0
    TileWallpaper_value = 1
ElseIf item.Key = "居中" Then
    WallpaperStyle_value = 0
    TileWallpaper_value = 0
End If
End Sub

'试听声音
Private Sub sound_Play_Click()
sndPlaySound url_to_N(url_sound), SND_ASYNC '播放一次，SND_ASYNC Or SND_LOOP '循环播放
End Sub
'停止试听声音
Private Sub sound_Stop_Click()
sndPlaySound vbNullString, SND_ASYNC
End Sub

'声音地址改变时
Private Sub url_sound_Change()
Dim i As Byte
'先判断新址是否是改变了的（白色的喇叭只有读取时一次性显示）
If url_sound <> Sound(2, sound_num) And url_sound <> "" Then '如果改变后的值和原来的不一样（且不为空）
        TreeView_Sound.Nodes("s" & sound_num).Image = 3 '将改变了值的节点小喇叭变黄
ElseIf url_sound.text = "" Then '如果为空
        TreeView_Sound.Nodes("s" & sound_num).Image = 0 '将改变了值的节点小喇叭删掉
End If

'将新地址储存进数组
Sound(2, sound_num) = url_sound.text
If url_sound <> "" And Dir(url_to_N(url_sound)) <> "" Then '不为空且文件存在
    sound_Play.Enabled = True
    sound_Stop.Enabled = True
Else
    sound_Play.Enabled = False
    sound_Stop.Enabled = False
End If
End Sub
'sound点击某个东西做什么
Private Sub TreeView_Sound_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer
If Left$(Node.Key, 1) <> "s" Then
    Command_sound.Enabled = False
    sound_Play.Enabled = False
    sound_Stop.Enabled = False
    url_sound.Enabled = False
    url_sound = ""
Else
    i = Mid$(Node.Key, 2) '从第二个开始就是编号
    sound_num = i
    If Check_snd.value = 0 Then
        Command_sound.Enabled = True
        url_sound.Enabled = True
    End If
    url_sound = Sound(2, sound_num)
End If
End Sub

'图标地址改变时
Private Sub url_icon_Change(Index As Integer)
Dim i As Byte
'将地址储存进数组
For i = 0 To 5
    SysIco(i, 0) = url_icon(i).text
Next
Draw_Ico '重画图标内容
End Sub

'选择图标文件
Private Sub Command_icon_Click(Index As Integer)
ico_num = Command_icon(Index).Index
'以前的老版只能打开单个
'On Error GoTo ErrHandler
'    CommonDialog1.DialogTitle = "选择图标文件"
'    CommonDialog1.Filter = "图标文件 (*.ico ; *.icon)|*.ico;*.icon|图标库文件 (*.exe ; *.dll)|*.exe;*.dll"
'    If url_icon(ico_num) <> "" Then '不为空
'        CommonDialog1.filename = url_to_N(url_icon(ico_num)) '打开时默认选择当前文件
'    Else
'        CommonDialog1.filename = ""
'    End If
'    CommonDialog1.ShowOpen
'    If InStr(CommonDialog1.filename, ".exe") <> 0 Or InStr(CommonDialog1.filename, ".dll") <> 0 Then
'        MsgBox "本程序暂不支持把图标库含有的文件都显示出来，然后再选择" + vbLf + "（不过将来可能支持）" + vbLf + "请手动输入图标编号", 64, "提示"
'        If InStr(CommonDialog1.filename, ".exe") <> 0 Then
'            url_icon(ico_num) = url_to_P(CommonDialog1.filename) & ",-0"
'        ElseIf InStr(CommonDialog1.filename, ".dll") <> 0 Then
'            url_icon(ico_num) = url_to_P(CommonDialog1.filename) & ",-0"
'        End If
'    Else
'        url_icon(ico_num) = url_to_P(CommonDialog1.filename)
'    End If
'
'Exit Sub
'ErrHandler:
''用户按“取消”按钮。
'Exit Sub

Dim file As String
Dim IconNum As Long
If InStr(url_icon(ico_num), ",") > 0 Then
    file = Left(url_icon(ico_num), InStr(url_icon(ico_num), ",") - 1)
    IconNum = Right(url_icon(ico_num), Len(url_icon(ico_num)) - InStr(url_icon(ico_num), ","))
Else
    file = url_icon(ico_num)
    IconNum = 0
End If

If chooseIcon(file, IconNum, Me) = True Then
    url_icon(ico_num) = file & "," & IconNum
End If

End Sub

'点击鼠标默认按钮
Private Sub cur_default_Click()
url_cur = SysCur(cur_num, 2)
End Sub

'选择鼠标文件
Private Sub Command_cur_Click()
On Error GoTo ErrHandler
    CommonDialog1.DialogTitle = Load_Lanuage("选择鼠标指针", "Public", "CommonDialog_Cur_DialogTitle_Load", Lanuage_Now)
    CommonDialog1.Filter = Load_Lanuage("光标", "Public", "CommonDialog_Cur_Filter", Lanuage_Now) & " (*.cur ; *.ani)|*.cur;*.ani"
    If url_cur <> "" Then '不为空
        CommonDialog1.filename = url_to_N(url_cur) '打开时默认选择当前文件
    Else
        CommonDialog1.filename = ""
    End If
    CommonDialog1.ShowOpen
    url_cur = url_to_P(CommonDialog1.filename)
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

'鼠标地址改变时
Private Sub url_cur_Change()
Dim i As Byte
'右上角动画改变
If url_cur <> "" And Dir(url_to_N(url_cur)) <> "" Then '不为空且文件存在
    cAni2.LoadFromFile url_to_N(url_cur)
Else
    cAni2.LoadFromFile ""
End If
'将地址储存进数组
    SysCur(cur_num, 0) = url_cur.text
Draw_Cur '重画鼠标列表预览 '重画鼠标列表预览
End Sub
'版本检查按钮
Private Sub Check_ver_Click()
Timer_Update.Enabled = True
End Sub



'点击每个鼠标选项时
Private Sub Cur_BG_Click(Index As Integer)
Dim i As Byte
cur_num = Cur_BG(Index).Index
url_cur = SysCur(cur_num, 0)
For i = 0 To 14
    If i = Main.Cur_BG(Index).Index Then
        Main.Cur_BG(i).BackColor = &H8000000D
        Main.Pic_Cur(i).BackColor = &H8000000D
        Main.Cur_BG(i).ForeColor = &H80000018
    Else
        If glass_ok = True Then
            Main.Cur_BG(i).BackColor = m_transparencyKey
            Main.Pic_Cur(i).BackColor = m_transparencyKey
        Else
            Main.Cur_BG(i).BackColor = &H80000005
            Main.Pic_Cur(i).BackColor = &H80000005
        End If
        Main.Cur_BG(i).ForeColor = &H80000012
    End If
Next
Draw_Cur '重画鼠标列表预览
End Sub

Private Sub Pic_Cur_Click(Index As Integer)
Dim i As Byte
cur_num = Pic_Cur(Index).Index
url_cur = SysCur(cur_num, 0)
For i = 0 To 14
    If i = Main.Cur_BG(Index).Index Then
        Main.Cur_BG(i).BackColor = &H8000000D
        Main.Pic_Cur(i).BackColor = &H8000000D
        Main.Cur_BG(i).ForeColor = &H80000018
    Else
        If glass_ok = True Then
            Main.Cur_BG(i).BackColor = m_transparencyKey
            Main.Pic_Cur(i).BackColor = m_transparencyKey
        Else
            Main.Cur_BG(i).BackColor = &H80000005
            Main.Pic_Cur(i).BackColor = &H80000005
        End If
        Main.Cur_BG(i).ForeColor = &H80000012
    End If
Next
Draw_Cur '重画鼠标列表预览
End Sub

Private Sub Value_System_Color_Change(Index As Integer)
'当在文本框按键修改时才存入
If change_System_color_text = True Then
    Dim i As Byte
    For i = 0 To 30
        SysColors(i, 1) = Me.Value_System_Color(i)
    Next
    If Main.ImageCombo_Classic_Style.ComboItems(1).Selected = False Then
        Main.ImageCombo_Classic_Style.ComboItems(1).Selected = True
    End If
End If
change_System_color_text = False
End Sub

Private Sub Value_System_Color_KeyPress(Index As Integer, KeyAscii As Integer)
    change_System_color_text = True '当在文本框按键才是修改
End Sub

'滚动条改变值控制鼠标内容
Private Sub VScroll_cur_Change()
    Frame_Mouse.Top = 0 - VScroll_cur.value * 5.9
End Sub
'滚动条滚动鼠标内容
Private Sub VScroll_cur_Scroll()
    Frame_Mouse.Top = 0 - VScroll_cur.value * 5.9
End Sub

'连续播放ANI动画
Private Sub Timer_cur_Timer()
If url_cur <> "" And Dir(url_to_N(url_cur)) <> "" Then '不为空且文件存在
    cAni2.Step
    cAni2.Draw ShowCur.hDC, 10, 10, , , ShowCur.BackColor
Else
    ShowCur.Cls
End If

End Sub
'优先于Load加载的前提代码
Private Sub Form_Initialize()
'MsgBox "您的操作系统是" & strOSversion & " " & System_Ver
    Call Change_Lanuage(Lanuage_Now)

End Sub
Private Sub Form_Load()

Dim i%

Call Creat_Default '调用模块里面建立默认值的列表★
Draw_Ico '画一次默认图标

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
'为了让鼠标选择框背景透明，不然就是白色，还有选择壁纸更换时间的那个下拉框
If glass_ok = True Then
    Mouse_box.BackColor = m_transparencyKey
    Frame_Mouse.BackColor = m_transparencyKey
    Frame_System_Color.BackColor = m_transparencyKey
    Frame_Edit_Panel_Tab.BackColor = m_transparencyKey
    Frame_select_mss.BackColor = m_transparencyKey
    System_Color_Tab_Frame.BackColor = m_transparencyKey
    mss_Aero.BackColor = m_transparencyKey
    Check_Alpha.BackColor = m_transparencyKey
    mss_Basic.BackColor = m_transparencyKey
    mss_Classic.BackColor = m_transparencyKey
    Check_insert_system_color.BackColor = m_transparencyKey
    Check_paper_change.BackColor = m_transparencyKey
    Check_snd.BackColor = m_transparencyKey
    Papers_Edit_Allow.BackColor = m_transparencyKey
    Frame_Main_Tab.BackColor = m_transparencyKey
    Command_Guide.BackColor = m_transparencyKey
Else
    Mouse_box.BackColor = &H80000005
    Frame_Mouse.BackColor = &H80000005
    Frame_System_Color.BackColor = &H8000000F
    Frame_Edit_Panel_Tab.BackColor = &H8000000F
    Frame_select_mss.BackColor = &H8000000F
    System_Color_Tab_Frame.BackColor = &H8000000F
    mss_Aero.BackColor = &H8000000F
    Check_Alpha.BackColor = &H8000000F
    mss_Basic.BackColor = &H8000000F
    mss_Classic.BackColor = &H8000000F
    Check_insert_system_color.BackColor = &H8000000F
    Check_paper_change.BackColor = &H8000000F
    Check_snd.BackColor = &H8000000F
    Papers_Edit_Allow.BackColor = &H8000000F
    Frame_Main_Tab.BackColor = &H8000000F
    Command_Guide.BackColor = &H8000000F
End If

'排列主面板
For i = 0 To Option_Main_Tab.UBound
    Main_Frame(i).BorderStyle = 0
    Main_Frame(i).Top = 72 * 15 '移动位置
    Main_Frame(i).Left = 0 * 15
    Main_Frame(i).Height = 372 * 15
    Main_Frame(i).Width = 497 * 15
    Main_Frame(i).Visible = False
    If glass_ok = True Then
        Main_Frame(i).BackColor = m_transparencyKey
        Option_Main_Tab(i).BackColor = m_transparencyKey
    Else
        Main_Frame(i).BackColor = &H8000000F
        Option_Main_Tab(i).BackColor = &H8000000F
    End If
Next
'排列编辑面板
For i = 0 To Edit_Panel_Frame.UBound
    Edit_Panel_Frame(i).Top = 0 * 15 '移动位置
    Edit_Panel_Frame(i).Left = 90 * 15
    Edit_Panel_Frame(i).Height = 5550
    Edit_Panel_Frame(i).Width = 6015
    Edit_Panel_Frame(i).Visible = False
    If glass_ok = True Then
        Edit_Panel_Frame(i).BackColor = m_transparencyKey
    Else
        Edit_Panel_Frame(i).BackColor = &H8000000F
    End If
Next
'排列颜色面板
For i = 0 To System_Color_Frame.UBound
    System_Color_Frame(i).Top = 2520 '移动位置
    System_Color_Frame(i).Left = 120
    If glass_ok = True Then
        System_Color_Frame(i).BackColor = m_transparencyKey
        System_Color_Tab(i).BackColor = m_transparencyKey
    Else
        System_Color_Frame(i).BackColor = &H8000000F
        System_Color_Tab(i).BackColor = &H8000000F
    End If
Next
Show_ColorizationColor.BackColor = RGB_To_BGR(Aplha_Back_Color)
Show_ColorizationAfterglow.BackColor = RGB_To_BGR(Aplha_Back_Color)
Main.Height = 513 * 15
Main.Width = 503 * 15

Edit_Panel_Tab(0).value = True
System_Color_Tab(0).value = True
Cur_BG_Click (0) '初始状态，鼠标选择第一个
End Sub
Private Sub Form_Paint()
If not_First_Load = False Then

'选择是否玻璃
'If MsgBox("是否以全玻璃窗口模式启动？" & vbCrLf & "注：全玻璃模式仅在Aero启用下有效，否则会杯具……", 292, "选择启动模式") = vbYes Then ' 如果用户单击No按钮，则停止Unload事件。
'    glass_ok = True
'Else
'    glass_ok = False
'End If
Exit_ok = True
not_First_Load = True
End If

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
'关于
Private Sub Command_about_Click()
If frmAbout.Visible = False Then
    frmAbout.Show 1 '加了1使其他的不能操作
Else
    Unload frmAbout
End If
End Sub
'一键取色工具
Private Sub Command_getcolor_Click()
If Get_color.Visible = False Then
    Get_color.Show
Else
    Unload Get_color
End If
End Sub
'退出按钮
Private Sub Command_exit_Click()
Dim a As Integer
        a = MsgBox(Load_Lanuage("本程序暂不会检测保存情况|现在退出之前未保存的工作将消失，您确定要退出吗？", "Main", "Exit_Warn", Lanuage_Now), 308, Load_Lanuage("警告", "Main", "Exit_Warn_Title", Lanuage_Now))
        If a = vbYes Then '当点确定后开始执行下面代码
            Unload Get_color
            End
        End If
End Sub
'选择视觉风格
Private Sub mss_Aero_Click()
Check_Alpha.Enabled = True
Check_Alpha.value = 1
url_mss.Enabled = True
Label_mss.Enabled = True
Command_mss.Enabled = True
End Sub

Private Sub mss_Basic_Click()
Check_Alpha.Enabled = False
Check_Alpha.value = 0
url_mss.Enabled = True
Label_mss.Enabled = True
Command_mss.Enabled = True
End Sub

Private Sub mss_Classic_Click()
Check_Alpha.Enabled = False
Check_Alpha.value = 0
url_mss.Enabled = False
Label_mss.Enabled = False
Command_mss.Enabled = False
End Sub
'切换玻璃窗口
Private Sub Change_glass_Click()
If glass_ok = True Then
    Exit_ok = False
    Unload Main
    Main.Show
ElseIf glass_ok = False Then
    Exit_ok = False
    Unload Main
    Main.Show
End If
End Sub
'右上角关闭按钮
Private Sub Form_Unload(Cancel As Integer)
If Exit_ok = True Then
    If MsgBox(Load_Lanuage("本程序暂不会检测保存情况|现在退出之前未保存的工作将消失，您确定要退出吗？", "Main", "Exit_Warn", Lanuage_Now), 308, Load_Lanuage("警告", "Main", "Exit_Warn_Title", Lanuage_Now)) = vbNo Then ' 如果用户单击No按钮，则停止Unload事件。
        Cancel = True
    Else
    End
    End If
ElseIf Exit_ok = False Then
    Exit_ok = True
    If MsgBox(Load_Lanuage("切换玻璃窗口暂不会保存您编辑的内容|现在退出之前未保存的工作将消失，您确定要切换吗？", "Main", "Change_Aero_Warn", Lanuage_Now), 308, Load_Lanuage("切换玻璃窗口", "Main", "Change_Aero_Warn_Title", Lanuage_Now)) = vbNo Then ' 如果用户单击No按钮，则停止Unload事件。
        Cancel = True
    Else
    glass_ok = Not glass_ok '切换当前是否玻璃了
    End If
End If
End Sub

'滚动条改变值控制系统颜色内容
Private Sub VScroll_System_color_Change()
Frame_System_Color.Top = 0 - VScroll_System_color.value * 7
End Sub
'滚动条滚动系统颜色内容
Private Sub VScroll_System_color_Scroll()
Frame_System_Color.Top = 0 - VScroll_System_color.value * 7
End Sub
