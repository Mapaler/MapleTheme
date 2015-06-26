Attribute VB_Name = "检测联网"
Option Explicit
  
  '调用外部默认浏览器打开网页的声明
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub CheckVer(ByRef Auto_Update As Boolean, ByVal From_form As Variant)
    Dim Ver() As String
    Dim xmlHTTP1 As Object
    Dim Dangqianbanben As String '定义变量存放当前版本
    Dim Zuixinbanben As String '定义变量存放最新版本
    Dim newbanben As String '放取到的网页
    Dim a As Integer
    Dim newver As Boolean
    Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP1.Open "get", CheckVer_Page, True
    xmlHTTP1.setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '清除缓存
    xmlHTTP1.Send
    Dim CiShu As Integer
    While xmlHTTP1.ReadyState <> 4
        DoEvents
    Wend
    Dangqianbanben = App.Major & "." & App.Minor & Beta & " Build " & App.Revision    '设置当前软件的版本
    newbanben = xmlHTTP1.responseText
    Erase Ver
    Ver = Split(xmlHTTP1.responseText, "|")
    newver = False '先设定为没有新版本
    If newbanben <> "" And InStr(newbanben, "错误") = 0 Then
        Zuixinbanben = Ver(0) & "." & Ver(1) & Ver(3) & " Build " & Ver(2)    '设置当前软件的版本
        If Ver(0) > App.Major Then '主版本
            newver = True
        Else
            If Ver(1) > App.Minor Then '次版本
                newver = True
            Else
                If Ver(2) > App.Revision Then '小版本
                    newver = True
                End If
            End If
        End If
        
        If newver = True Then
            a = MsgBox(Load_Lanuage("测到有新的版本", "Public", "Found_New_Ver_Text1", Lanuage_Now) & vbCrLf & vbCrLf & Load_Lanuage("您的当前版本：", "Public", "Found_New_Ver_Text2", Lanuage_Now) & Dangqianbanben _
            & vbCrLf & Load_Lanuage("最新版本：", "Public", "Found_New_Ver_Text3", Lanuage_Now) & Zuixinbanben & vbCrLf & Load_Lanuage("更新日志：", "Public", "Found_New_Ver_Text5", Lanuage_Now) & vbCrLf & GetLog() & vbCrLf & vbCrLf & Load_Lanuage("立即更新吗？", "Public", "Found_New_Ver_Text4", Lanuage_Now), 68, Load_Lanuage("发现新版本", "Public", "Found_New_Ver_Title", Lanuage_Now))
            If a = 6 Then '当点确定后开始执行下面代码
                ShellExecute From_form.hWnd, vbNullString, WebSite, vbNullString, vbNullString, SW_SHOWNORMAL
            End If
        Else
            If Auto_Update = False Then
                a = MsgBox(Load_Lanuage("未发现新版本", "Public", "No_Found_New_Ver_Text1", Lanuage_Now) & vbCrLf & vbCrLf & Load_Lanuage("您的当前版本：", "Public", "No_Found_New_Ver_Text2", Lanuage_Now) & Dangqianbanben _
                & vbCrLf & Load_Lanuage("最新版本：", "Public", "No_Found_New_Ver_Text3", Lanuage_Now) & Zuixinbanben, vbOKOnly, Load_Lanuage("未发现新版本", "Public", "No_Found_New_Ver_Title", Lanuage_Now))
            End If
        End If
    Else
        If Auto_Update = False Then
            MsgBox Load_Lanuage("获取最新版本号失败，请检查是否连接上Internet，也有可能是服务器不正常。", "Public", "No_Connect_Text", Lanuage_Now), 64, Load_Lanuage("连接失败", "Public", "No_Connect_Title", Lanuage_Now)
        End If
    End If

    Set xmlHTTP1 = Nothing
End Sub

Public Function GetLog() As String
    Dim xmlHTTP2 As Object
    Dim log As String
    Set xmlHTTP2 = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP2.Open "get", Log_Page, True
    xmlHTTP2.setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '清除缓存
    xmlHTTP2.Send
    Dim CiShu As Integer
    While xmlHTTP2.ReadyState <> 4
        DoEvents
    Wend
    log = xmlHTTP2.responseText
    GetLog = log
    Set xmlHTTP2 = Nothing
End Function

