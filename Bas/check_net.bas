Attribute VB_Name = "�������"
Option Explicit
  
  '�����ⲿĬ�����������ҳ������
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub CheckVer(ByRef Auto_Update As Boolean, ByVal From_form As Variant)
    Dim Ver() As String
    Dim xmlHTTP1 As Object
    Dim Dangqianbanben As String '���������ŵ�ǰ�汾
    Dim Zuixinbanben As String '�������������°汾
    Dim newbanben As String '��ȡ������ҳ
    Dim a As Integer
    Dim newver As Boolean
    Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP1.Open "get", CheckVer_Page, True
    xmlHTTP1.setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '�������
    xmlHTTP1.Send
    Dim CiShu As Integer
    While xmlHTTP1.ReadyState <> 4
        DoEvents
    Wend
    Dangqianbanben = App.Major & "." & App.Minor & Beta & " Build " & App.Revision    '���õ�ǰ����İ汾
    newbanben = xmlHTTP1.responseText
    Erase Ver
    Ver = Split(xmlHTTP1.responseText, "|")
    newver = False '���趨Ϊû���°汾
    If newbanben <> "" And InStr(newbanben, "����") = 0 Then
        Zuixinbanben = Ver(0) & "." & Ver(1) & Ver(3) & " Build " & Ver(2)    '���õ�ǰ����İ汾
        If Ver(0) > App.Major Then '���汾
            newver = True
        Else
            If Ver(1) > App.Minor Then '�ΰ汾
                newver = True
            Else
                If Ver(2) > App.Revision Then 'С�汾
                    newver = True
                End If
            End If
        End If
        
        If newver = True Then
            a = MsgBox(Load_Lanuage("�⵽���µİ汾", "Public", "Found_New_Ver_Text1", Lanuage_Now) & vbCrLf & vbCrLf & Load_Lanuage("���ĵ�ǰ�汾��", "Public", "Found_New_Ver_Text2", Lanuage_Now) & Dangqianbanben _
            & vbCrLf & Load_Lanuage("���°汾��", "Public", "Found_New_Ver_Text3", Lanuage_Now) & Zuixinbanben & vbCrLf & Load_Lanuage("������־��", "Public", "Found_New_Ver_Text5", Lanuage_Now) & vbCrLf & GetLog() & vbCrLf & vbCrLf & Load_Lanuage("����������", "Public", "Found_New_Ver_Text4", Lanuage_Now), 68, Load_Lanuage("�����°汾", "Public", "Found_New_Ver_Title", Lanuage_Now))
            If a = 6 Then '����ȷ����ʼִ���������
                ShellExecute From_form.hWnd, vbNullString, WebSite, vbNullString, vbNullString, SW_SHOWNORMAL
            End If
        Else
            If Auto_Update = False Then
                a = MsgBox(Load_Lanuage("δ�����°汾", "Public", "No_Found_New_Ver_Text1", Lanuage_Now) & vbCrLf & vbCrLf & Load_Lanuage("���ĵ�ǰ�汾��", "Public", "No_Found_New_Ver_Text2", Lanuage_Now) & Dangqianbanben _
                & vbCrLf & Load_Lanuage("���°汾��", "Public", "No_Found_New_Ver_Text3", Lanuage_Now) & Zuixinbanben, vbOKOnly, Load_Lanuage("δ�����°汾", "Public", "No_Found_New_Ver_Title", Lanuage_Now))
            End If
        End If
    Else
        If Auto_Update = False Then
            MsgBox Load_Lanuage("��ȡ���°汾��ʧ�ܣ������Ƿ�������Internet��Ҳ�п����Ƿ�������������", "Public", "No_Connect_Text", Lanuage_Now), 64, Load_Lanuage("����ʧ��", "Public", "No_Connect_Title", Lanuage_Now)
        End If
    End If

    Set xmlHTTP1 = Nothing
End Sub

Public Function GetLog() As String
    Dim xmlHTTP2 As Object
    Dim log As String
    Set xmlHTTP2 = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP2.Open "get", Log_Page, True
    xmlHTTP2.setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '�������
    xmlHTTP2.Send
    Dim CiShu As Integer
    While xmlHTTP2.ReadyState <> 4
        DoEvents
    Wend
    log = xmlHTTP2.responseText
    GetLog = log
    Set xmlHTTP2 = Nothing
End Function

