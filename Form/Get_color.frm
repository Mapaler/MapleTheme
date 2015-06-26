VERSION 5.00
Begin VB.Form Get_color 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "一键取色工具"
   ClientHeight    =   6015
   ClientLeft      =   13470
   ClientTop       =   3630
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.PictureBox System_Color_box 
      Height          =   3495
      Left            =   3360
      ScaleHeight     =   3435
      ScaleWidth      =   3315
      TabIndex        =   29
      Top             =   120
      Width           =   3375
      Begin VB.VScrollBar VScroll_System_color 
         Height          =   3435
         LargeChange     =   100
         Left            =   3060
         Max             =   1000
         SmallChange     =   50
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame Frame_System_Color 
         BorderStyle     =   0  'None
         Caption         =   "系统颜色"
         Height          =   4215
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3135
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   0
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   62
            Text            =   "255 255 255"
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   1
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   61
            Text            =   "255 255 255"
            Top             =   340
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   2
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   60
            Text            =   "255 255 255"
            Top             =   580
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   3
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   59
            Text            =   "255 255 255"
            Top             =   925
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   4
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   58
            Text            =   "255 255 255"
            Top             =   1180
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   5
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   57
            Text            =   "255 255 255"
            Top             =   1525
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   6
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   56
            Text            =   "255 255 255"
            Top             =   1765
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   7
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   55
            Text            =   "255 255 255"
            Top             =   2110
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   8
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   54
            Text            =   "255 255 255"
            Top             =   2380
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   9
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   53
            Text            =   "255 255 255"
            Top             =   2725
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   10
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   52
            Text            =   "255 255 255"
            Top             =   2965
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   11
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   51
            Text            =   "255 255 255"
            Top             =   3310
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   12
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   50
            Text            =   "255 255 255"
            Top             =   3565
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   13
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   49
            Text            =   "255 255 255"
            Top             =   3910
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   14
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   48
            Text            =   "255 255 255"
            Top             =   4150
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   15
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   47
            Text            =   "255 255 255"
            Top             =   4495
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   16
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   46
            Text            =   "255 255 255"
            Top             =   4780
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   17
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   45
            Text            =   "255 255 255"
            Top             =   5125
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   18
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   44
            Text            =   "255 255 255"
            Top             =   5365
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   19
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   43
            Text            =   "255 255 255"
            Top             =   5710
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   20
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   42
            Text            =   "255 255 255"
            Top             =   5965
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   21
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   41
            Text            =   "255 255 255"
            Top             =   6310
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   22
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   40
            Text            =   "255 255 255"
            Top             =   6550
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   23
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   39
            Text            =   "255 255 255"
            Top             =   6895
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   24
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   38
            Text            =   "255 255 255"
            Top             =   7180
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   25
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   37
            Text            =   "255 255 255"
            Top             =   7525
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   26
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   36
            Text            =   "255 255 255"
            Top             =   7765
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   27
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   35
            Text            =   "255 255 255"
            Top             =   8110
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   28
            Left            =   1970
            MaxLength       =   11
            TabIndex        =   34
            Text            =   "255 255 255"
            Top             =   8365
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   29
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   33
            Text            =   "255 255 255"
            Top             =   8710
            Width           =   1095
         End
         Begin VB.TextBox Value_System_Color 
            Height          =   270
            Index           =   30
            Left            =   1965
            MaxLength       =   11
            TabIndex        =   32
            Text            =   "255 255 255"
            Top             =   8950
            Width           =   1095
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   93
            Top             =   20
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   92
            Top             =   355
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   2
            Left            =   45
            TabIndex        =   91
            Top             =   595
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   90
            Top             =   940
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   4
            Left            =   45
            TabIndex        =   89
            Top             =   1195
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   88
            Top             =   1540
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   87
            Top             =   1780
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   86
            Top             =   2125
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   8
            Left            =   45
            TabIndex        =   85
            Top             =   2395
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   84
            Top             =   2740
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   10
            Left            =   45
            TabIndex        =   83
            Top             =   2980
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   82
            Top             =   3325
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   12
            Left            =   45
            TabIndex        =   81
            Top             =   3580
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   80
            Top             =   3925
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   14
            Left            =   45
            TabIndex        =   79
            Top             =   4165
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   78
            Top             =   4510
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   16
            Left            =   45
            TabIndex        =   77
            Top             =   4795
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   76
            Top             =   5140
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   18
            Left            =   45
            TabIndex        =   75
            Top             =   5380
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   74
            Top             =   5725
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   20
            Left            =   45
            TabIndex        =   73
            Top             =   5980
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   72
            Top             =   6325
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   22
            Left            =   45
            TabIndex        =   71
            Top             =   6565
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   70
            Top             =   6910
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   24
            Left            =   45
            TabIndex        =   69
            Top             =   7195
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   68
            Top             =   7540
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   26
            Left            =   45
            TabIndex        =   67
            Top             =   7780
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   66
            Top             =   8125
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   28
            Left            =   45
            TabIndex        =   65
            Top             =   8380
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   64
            Top             =   8725
            Width           =   1935
         End
         Begin VB.Label Lable_System_Color 
            BackStyle       =   0  'Transparent
            Caption         =   "GradientInactiveTitle"
            Height          =   255
            Index           =   30
            Left            =   45
            TabIndex        =   63
            Top             =   8965
            Width           =   1935
         End
      End
   End
   Begin VB.CommandButton Command_exit 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command_window 
      Caption         =   "打开窗体颜色和外观面板"
      Height          =   495
      Left            =   3360
      TabIndex        =   27
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Command_glass 
      Caption         =   "打开透明颜色面板"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton freshen2 
      Caption         =   "刷新"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton add_all2 
      Caption         =   "插入全部"
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command_ColorizationBlurBalance 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command_ColorizationGlassReflectionIntensity 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command_ColorizationAfterglowBalance 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command_ColorizationAfterglow 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command_ColorizationColorBalance 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command_ColorizationColor 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command_mss 
      Caption         =   "插入"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Value_ColorizationAfterglow 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Value_ColorizationColor 
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Value_ColorizationColorBalance 
      Height          =   270
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Value_ColorizationBlurBalance 
      Height          =   270
      Left            =   1560
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Value_ColorizationAfterglowBalance 
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Value_ColorizationGlassReflectionIntensity 
      Height          =   270
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox url_mss 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton add_all 
      Caption         =   "插入全部"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton freshen 
      Caption         =   "刷新"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label_help 
      BackStyle       =   0  'Transparent
      Caption         =   "说明"
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   6615
   End
   Begin VB.Label Label_mss 
      BackStyle       =   0  'Transparent
      Caption         =   "视觉风格文件"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label_ColorizationGlassReflectionIntensity 
      BackStyle       =   0  'Transparent
      Caption         =   "Aero条纹数量"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "大背景透明度"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label_ColorizationBlurBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "模糊平衡"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label_ColorizationAfterglowBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "发光颜色平衡"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label_ColorizationColorBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "主颜色平衡"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label_ColorizationAfterglow 
      BackStyle       =   0  'Transparent
      Caption         =   "发光颜色"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label_ColorizationColor 
      BackStyle       =   0  'Transparent
      Caption         =   "主颜色"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Get_color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_all_Click()
Main.Value_ColorizationAfterglow = Me.Value_ColorizationAfterglow
Main.Value_ColorizationAfterglowBalance = Me.Value_ColorizationAfterglowBalance
Main.Value_ColorizationBlurBalance = Me.Value_ColorizationBlurBalance
Main.Value_ColorizationColor = Me.Value_ColorizationColor
Main.Value_ColorizationColorBalance = Me.Value_ColorizationColorBalance
Main.Value_ColorizationGlassReflectionIntensity = Me.Value_ColorizationGlassReflectionIntensity
Main.url_mss = Me.url_mss
End Sub

Private Sub add_all2_Click()
For i = 0 To 30
SysColors(i, 1) = Me.Value_System_Color(i)
Next
Main.ImageCombo_Classic_Style.ComboItems(1).Selected = True
End Sub

Private Sub Command_ColorizationAfterglow_Click()
Main.Value_ColorizationAfterglow = Me.Value_ColorizationAfterglow
End Sub

Private Sub Command_ColorizationAfterglowBalance_Click()
Main.Value_ColorizationAfterglowBalance = Me.Value_ColorizationAfterglowBalance
End Sub

Private Sub Command_ColorizationBlurBalance_Click()
Main.Value_ColorizationBlurBalance = Me.Value_ColorizationBlurBalance
End Sub

Private Sub Command_ColorizationColor_Click()
Main.Value_ColorizationColor = Me.Value_ColorizationColor
End Sub

Private Sub Command_ColorizationColorBalance_Click()
Main.Value_ColorizationColorBalance = Me.Value_ColorizationColorBalance
End Sub

Private Sub Command_ColorizationGlassReflectionIntensity_Click()
Main.Value_ColorizationGlassReflectionIntensity = Me.Value_ColorizationGlassReflectionIntensity
End Sub

Private Sub Command_exit_Click()
    Unload Me
End Sub

Private Sub Command_glass_Click()
If System_Ver < 6 Then
    MsgBox Load_Lanuage("检测到您的系统为", "Main", "My_System1", Lanuage_Now) & strOSversion & Load_Lanuage("，您的操作系统版本并无此功能。", "Main", "My_System2", Lanuage_Now)
Else
    Call Shell("control.exe /name Microsoft.Personalization /page pageColorization") '修改透明颜色
End If
End Sub

Private Sub Command_mss_Click()
Main.url_mss = Me.url_mss
End Sub

Private Sub Command_window_Click()
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,advanced,@advanced") '打开窗体设置和外观
End Sub

Private Sub Form_Load()
    Call Change_Lanuage(Lanuage_Now)

    Me.Icon = Main.Icon
    Me.Left = Main.Left + Main.Width
    Me.Top = Main.Top + 1000
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '保持在前
'Label_help.Caption = "使用方法：" & vbCrLf & "先确保是在Aero风格下（Win7 HomeBasic下叫做Windows 7 Standard），使用Windows自带的个性化或者魔方（Aero效果调节）调节好您所满意的颜色，并保存" _
'& vbCrLf & "然后点击本工具的“刷新”，就会去读您当前的设置值，然后根据需要选择插入到主程序窗口里面去。" & vbCrLf & "右边则是系统基本颜色设置，非调节Classic不用使用。"

'对系统颜色们进行排列
For i = 0 To 30
    Lable_System_Color(i).Caption = SysColors(i, 0) '设置左边的名称
    Lable_System_Color(i).Top = 290 * i + 60
    Lable_System_Color(i).Left = 50
    Value_System_Color(i).Top = 290 * i + 45
    Value_System_Color(i).Left = 1970
    Value_System_Color(i) = GetString(HKEY_CURRENT_USER, "Control Panel\Colors", SysColors(i, 0))
Next
Frame_System_Color.Width = 209 * 15
Frame_System_Color.Height = 900 * 15

'刷新值
url_mss = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName")
Value_ColorizationColor = "0x" & x10_to_x16(GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColor"), 8)
Value_ColorizationColorBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColorBalance")
Value_ColorizationAfterglow = "0x" & x10_to_x16(GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglow"), 8)
Value_ColorizationAfterglowBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglowBalance")
Value_ColorizationGlassReflectionIntensity = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationGlassReflectionIntensity")
Value_ColorizationBlurBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationBlurBalance")
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
    MsgBoxErr.description
'全玻璃↑
End If
If glass_ok = True Then
    Frame_System_Color.BackColor = m_transparencyKey
Else
    Frame_System_Color.BackColor = &H8000000F
End If
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

'刷新值
Private Sub freshen_Click()
url_mss = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName")
Value_ColorizationColor = "0x" & x10_to_x16(GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColor"), 8)
Value_ColorizationColorBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationColorBalance")
Value_ColorizationAfterglow = "0x" & x10_to_x16(GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglow"), 8)
Value_ColorizationAfterglowBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationAfterglowBalance")
Value_ColorizationGlassReflectionIntensity = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationGlassReflectionIntensity")
Value_ColorizationBlurBalance = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", "ColorizationBlurBalance")
End Sub

Private Sub freshen2_Click()
Dim i As Byte
For i = 0 To 30
    Value_System_Color(i) = GetString(HKEY_CURRENT_USER, "Control Panel\Colors", SysColors(i, 0))
Next
End Sub

'滚动条改变值控制系统颜色内容
Private Sub VScroll_System_color_Change()
    Frame_System_Color.Top = 0 - VScroll_System_color.value * 5.8
End Sub
'滚动条滚动系统颜色内容
Private Sub VScroll_System_color_Scroll()
    Frame_System_Color.Top = 0 - VScroll_System_color.value * 5.8
End Sub
