Attribute VB_Name = "全局数组"
Public Const App_Beta = "大学毕业版" '本程序的当前Beta等版本名
Public Const pWidth = 100 '壁纸预览宽
Public Const pHeight = 75 '壁纸预览高
'Public Const CheckVer_Page = "http://www.mapaler.com/tools/checkver/?mod=checkver&soft=mapletheme" '最新版本号
'Public Const Log_Page = "http://www.mapaler.com/tools/checkver/?soft=mapletheme&from=soft" '软件更新LOG
Public Const UpdataURL = "http://www.mapaler.com/tools/checkver/index.php?soft=mapletheme&mod=xmlverinfo&ver=newest"
Public Const WebSite = "http://www.mapaler.com/mapletheme" '官网地址

Global Auto_Update As Boolean '判断是自动更新还是点击更新按钮

Global BAT_Color(15) As Long '16色
'BAT的16色

Global Config_Url As String

'Global Theme1() As String '创建主题列表数组。
'Global Theme2() As String '创建主题列表数组。
Global Theme1 As Collection '创建主题列表数组。
Global Theme2 As Collection '创建主题列表数组。

Global SystemTextShow As Boolean '声音等项目是否从文件读取
Global Sound_Style As Integer '声音列表风格
Global F_Sound() As String '音效父节点
Global Sound() As String '创建系统音效项名称与储存位置的二维数组。
'0：英文，1：中文，2：当前文件地址
Global Sound_Name() As String '创建系统音效名称的数组。

Global SysCur(14, 100) As String '创建鼠标地址数组。
'0：当前文件地址，1：英文名称，2：默认地址,3：中文名称
Global SysIco(6, 3) As String '创建图标地址数组。
'0：当前文件地址，1：英文名称，2：中文名称，3：默认地址
Global SysColors(31, 6) As String '创建系统默认颜色。
'0：英文名，1：当前颜色值，2：Windows 经典，3：高对比度 #1，4：高对比度 #2，5：高对比度黑色，6：高对比度白色

Public Const m_transparencyKey As Long = &HFEFFFF   '全玻璃扣掉的颜色
Global glass_ok As Boolean '是否开玻璃

Global Aplha_Back_Color As Long '透明色背景色（伪）
Global Exit_ok As Boolean   '是否是退出

Global TileWallpaper_value As Byte, WallpaperStyle_value As Byte '桌面壁纸的显示方式
'Global PaperFileName() As String '壁纸名的数组
Global PaperFileName As Collection '壁纸名的数组
Global System_Ver As Single  '系统版本
Global strOSversion As String '系统版本(文字)
'Global Lanuages() As String '语言文件的数组
Global Lanuages As Collection '语言文件的数组
Global Lanuage_Now As String '当前语言
Public Const Lanuage_Need = "2.4.60" '需要语言版本
Global SysRoot As String '操作系统文件夹路径
Global SysPath As Byte '转换成什么环境变量

Global AutoPaper As Byte '是否自动传送图片列表
Global New_List As String '传递的壁纸列表

'正则表达式
Public Const FileURL_Parten = "([A-Za-z]:|\\)[\\/][^:\*\?""<>\|]*" '真实文件路径正则表达式
