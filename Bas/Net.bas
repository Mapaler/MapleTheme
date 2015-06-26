Attribute VB_Name = "��������_������"
Option Explicit
  
  '�����ⲿĬ�����������ҳ������
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Type aSoftVerInfo
    tName As String
    tShortname As String
    tWebsite As String
    tMajor As Long
    tMinor As Long
    tRevision As Long
    tBeta As String
    tTime As String
    tLog As String
End Type
'��ȡ��ҳ����
Public Function GetWebText(ByVal tURL As String)
    Dim xmlHTTP1 As Object
    Set xmlHTTP1 = CreateObject("MSXML2.XMLHTTP")
    xmlHTTP1.Open "get", tURL, True
    xmlHTTP1.setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '�������
    xmlHTTP1.send
    Dim CiShu As Integer
    While xmlHTTP1.readyState <> 4
        DoEvents
    Wend
    GetWebText = xmlHTTP1.responseText
End Function
'��XML���ݴ�������汾��Ϣ����
Public Function GetSoftVersionFromXML(XMLinfo As DOMDocument) As aSoftVerInfo
    Dim tVerInfo As aSoftVerInfo
    
    tVerInfo.tName = GetNodeAttribute(XMLinfo, "/versioninfo/soft", "name")
    tVerInfo.tShortname = GetNodeAttribute(XMLinfo, "/versioninfo/soft", "shortname")
    
    Dim VerTemp() As String
    Erase VerTemp
    VerTemp = Split(GetNodeAttribute(XMLinfo, "/versioninfo/soft", "version", "0.0.0"), ".")
    tVerInfo.tMajor = VerTemp(0)
    tVerInfo.tMinor = VerTemp(1)
    tVerInfo.tRevision = VerTemp(2)
    
    tVerInfo.tBeta = GetNodeAttribute(XMLinfo, "/versioninfo/soft", "beta")
    
    tVerInfo.tTime = GetNodeAttribute(XMLinfo, "/versioninfo/soft", "time", "0000-00-00")
    tVerInfo.tWebsite = GetNodeAttribute(XMLinfo, "/versioninfo/soft", "website")
    
    tVerInfo.tLog = GetNodeText(XMLinfo, "/versioninfo/log", "")
    GetSoftVersionFromXML = tVerInfo
End Function
'����ҳlog���VB�ĸ�ʽ
Public Function ChangeLogTextToVB(ByVal yText As String) As String
    Dim TextTemp_1 As String, TextTemp_2 As String
    Dim TextTempPV_1() As PatternValue, TextTempPV_2() As PatternValue
    Dim NumTemp_1 As Long, NumTemp_2 As Long
    Dim i As Long, j As Long
    TextTemp_1 = yText
    '�滻�б���
    NumTemp_1 = SearchText(TextTemp_1, "<ul[^>]*>([\d\D]*?)<\/ul[^>]*>", TextTempPV_1())
    For i = 0 To NumTemp_1 - 1
    
        NumTemp_2 = SearchText(TextTempPV_1(i).InValue(0), "<li[^>]*?>([^<]*)", TextTempPV_2())
        TextTemp_2 = ""
        For j = 0 To NumTemp_2 - 1
            TextTemp_2 = TextTemp_2 & j + 1 & "��" & TextTempPV_2(j).InValue(0) & vbCrLf
        Next j
        TextTemp_1 = Replace(TextTemp_1, TextTempPV_1(i).AllValue, TextTemp_2) '���ﲻ��������ʽ
    Next i
    '�滻���в���
    TextTemp_1 = ReplaceText(TextTemp_1, "<[^br<>]*br[^br<>]*>", vbCrLf)
    ChangeLogTextToVB = TextTemp_1
End Function
Public Sub CheckVer(ByVal tUpdataURL As String, ByRef Auto_Update As Boolean, ByVal From_form As Variant)
    Dim UpdataInfoXML As DOMDocument '����XML����
    Dim newVer As aSoftVerInfo
    Dim CheckError As Boolean
    Dim TextTemp_1 As String, TextTemp_2 As String
    Set UpdataInfoXML = New DOMDocument
    '������ҳԴ����
    Dim TextTemp As String
    TextTemp = GetWebText(tUpdataURL)
    
    If TextTemp = "" Or InStr(TextTemp, "Error") = 1 Then
        '��ҳû�л����Ǵ�����Ǵ����
        CheckError = True
    End If
    
    '�����XML
    UpdataInfoXML.loadXML TextTemp
    If UpdataInfoXML.documentElement Is Nothing Then
        '����XML�����ļ�ʧ��
        CheckError = True
    End If
    
If CheckError = True Then
    If Auto_Update = False Then
        TextTemp_1 = "��ȡ���°汾��ʧ�ܣ������Ƿ�������Internet��Ҳ�п����Ƿ�������������"
        TextTemp_2 = "����ʧ��"
        MsgBox TextTemp_1, vbCritical, TextTemp_2
    End If
Else
    '����XML���ݽ�����
    newVer = GetSoftVersionFromXML(UpdataInfoXML)
    
    Dim aMsg As Long
    Dim Dangqianbanben As String
    Dim Zuixinbanben As String
    Dim LogText_VB As String
    Dangqianbanben = App.Major & "." & App.Minor & "." & App.Revision  '����ĵ�ǰ�汾
    If App_Beta <> "" Then
        Dangqianbanben = Dangqianbanben & " " & App_Beta
    End If
    Zuixinbanben = newVer.tMajor & "." & newVer.tMinor & "." & newVer.tRevision '��������°汾
    If newVer.tBeta <> "" Then
        Zuixinbanben = Zuixinbanben & " " & newVer.tBeta
    End If
    LogText_VB = ChangeLogTextToVB(newVer.tLog)
    If newVer.tMajor > App.Major Or _
    (newVer.tMajor = App.Major And newVer.tMinor > App.Minor) Or _
    (newVer.tMajor = App.Major And newVer.tMinor = App.Minor And newVer.tRevision > App.Revision) Then
    
        TextTemp_1 = " ��⵽���µİ汾" & vbCrLf
        TextTemp_1 = TextTemp_1 & "����ʱ��" & newVer.tTime & vbCrLf & vbCrLf
        TextTemp_1 = TextTemp_1 & "���ĵ�ǰ�汾��" & Dangqianbanben & vbCrLf
        TextTemp_1 = TextTemp_1 & "���°汾��    " & Zuixinbanben & vbCrLf
        TextTemp_1 = TextTemp_1 & "������־��" & vbCrLf
        TextTemp_1 = TextTemp_1 & LogText_VB & vbCrLf & vbCrLf
        TextTemp_1 = TextTemp_1 & "����������"
        TextTemp_2 = newVer.tName & "�����°汾"
        
        aMsg = MsgBox(TextTemp_1, vbYesNo Or vbInformation, TextTemp_2)
        If aMsg = vbYes Then '����ȷ����ʼִ���������
            ShellExecute From_form.hwnd, vbNullString, newVer.tWebsite, vbNullString, vbNullString, SW_SHOWNORMAL
        End If
    Else
        If Auto_Update = False Then
            
        TextTemp_1 = "δ�����°汾" & vbCrLf & vbCrLf
        TextTemp_1 = TextTemp_1 & "���ĵ�ǰ�汾��" & Dangqianbanben & vbCrLf
        TextTemp_1 = TextTemp_1 & "���°汾��" & Zuixinbanben & vbCrLf & vbCrLf
        TextTemp_1 = TextTemp_1 & "�Ƿ������վ�鿴��"
        TextTemp_2 = newVer.tName & "δ�����°汾"
        
            aMsg = MsgBox(TextTemp_1, vbYesNo, TextTemp_2)
            If aMsg = vbYes Then '����ȷ����ʼִ���������
                ShellExecute From_form.hwnd, vbNullString, newVer.tWebsite, vbNullString, vbNullString, SW_SHOWNORMAL
            End If
        End If
    End If
End If
End Sub
