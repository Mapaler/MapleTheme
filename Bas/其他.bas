Attribute VB_Name = "ȫ������"
Public Const App_Beta = "��ѧ��ҵ��" '������ĵ�ǰBeta�Ȱ汾��
Public Const pWidth = 100 '��ֽԤ����
Public Const pHeight = 75 '��ֽԤ����
'Public Const CheckVer_Page = "http://www.mapaler.com/tools/checkver/?mod=checkver&soft=mapletheme" '���°汾��
'Public Const Log_Page = "http://www.mapaler.com/tools/checkver/?soft=mapletheme&from=soft" '�������LOG
Public Const UpdataURL = "http://www.mapaler.com/tools/checkver/index.php?soft=mapletheme&mod=xmlverinfo&ver=newest"
Public Const WebSite = "http://www.mapaler.com/mapletheme" '������ַ

Global Auto_Update As Boolean '�ж����Զ����»��ǵ�����°�ť

Global BAT_Color(15) As Long '16ɫ
'BAT��16ɫ

Global Config_Url As String

'Global Theme1() As String '���������б����顣
'Global Theme2() As String '���������б����顣
Global Theme1 As Collection '���������б����顣
Global Theme2 As Collection '���������б����顣

Global SystemTextShow As Boolean '��������Ŀ�Ƿ���ļ���ȡ
Global Sound_Style As Integer '�����б���
Global F_Sound() As String '��Ч���ڵ�
Global Sound() As String '����ϵͳ��Ч�������봢��λ�õĶ�ά���顣
'0��Ӣ�ģ�1�����ģ�2����ǰ�ļ���ַ
Global Sound_Name() As String '����ϵͳ��Ч���Ƶ����顣

Global SysCur(14, 100) As String '��������ַ���顣
'0����ǰ�ļ���ַ��1��Ӣ�����ƣ�2��Ĭ�ϵ�ַ,3����������
Global SysIco(6, 3) As String '����ͼ���ַ���顣
'0����ǰ�ļ���ַ��1��Ӣ�����ƣ�2���������ƣ�3��Ĭ�ϵ�ַ
Global SysColors(31, 6) As String '����ϵͳĬ����ɫ��
'0��Ӣ������1����ǰ��ɫֵ��2��Windows ���䣬3���߶Աȶ� #1��4���߶Աȶ� #2��5���߶ԱȶȺ�ɫ��6���߶ԱȶȰ�ɫ

Public Const m_transparencyKey As Long = &HFEFFFF   'ȫ�����۵�����ɫ
Global glass_ok As Boolean '�Ƿ񿪲���

Global Aplha_Back_Color As Long '͸��ɫ����ɫ��α��
Global Exit_ok As Boolean   '�Ƿ����˳�

Global TileWallpaper_value As Byte, WallpaperStyle_value As Byte '�����ֽ����ʾ��ʽ
'Global PaperFileName() As String '��ֽ��������
Global PaperFileName As Collection '��ֽ��������
Global System_Ver As Single  'ϵͳ�汾
Global strOSversion As String 'ϵͳ�汾(����)
'Global Lanuages() As String '�����ļ�������
Global Lanuages As Collection '�����ļ�������
Global Lanuage_Now As String '��ǰ����
Public Const Lanuage_Need = "2.4.60" '��Ҫ���԰汾
Global SysRoot As String '����ϵͳ�ļ���·��
Global SysPath As Byte 'ת����ʲô��������

Global AutoPaper As Byte '�Ƿ��Զ�����ͼƬ�б�
Global New_List As String '���ݵı�ֽ�б�

'������ʽ
Public Const FileURL_Parten = "([A-Za-z]:|\\)[\\/][^:\*\?""<>\|]*" '��ʵ�ļ�·��������ʽ
