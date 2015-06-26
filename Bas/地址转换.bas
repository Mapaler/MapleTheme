Attribute VB_Name = "��ַת��"
Option Explicit

'�������������⻷��������ַת��
Public Function url_to_S(ByVal url_old As String) As String
Dim SysRoot_From_Sys As String '��ȡ�Ļ������� ����ϵͳ�ļ���·��
    If SysRoot <> "" Then
        If SysRoot = "0" Then
            SysRoot_From_Sys = Environ("systemroot")
        Else
            SysRoot_From_Sys = SysRoot
        End If
    Else
        SysRoot_From_Sys = Environ("systemroot")
    End If
    
Dim SysPath_Choose As String
        If SysPath = 0 Then
            SysPath_Choose = "%SystemRoot%"
        ElseIf SysPath = 1 Then
            SysPath_Choose = "%WinDir%"
        ElseIf SysPath = 2 Then
            SysPath_Choose = SysRoot_From_Sys
        Else
            SysPath_Choose = "%SystemRoot%"
        End If
        
Dim url_now As String
    If InStr(url_old, "%ResourceDir%") <> 0 Then '�ں�ţ�����﷢�ֵġ���
        url_now = SysPath_Choose & "\Resources" + Mid(url_old, InStr(url_old, "%ResourceDir%") + Len("%ResourceDir%"))
    Else
        url_now = url_old
    End If
    url_to_S = url_now
End Function

'����ͨ��ַת���ɻ���������ַ
Public Function url_to_P(ByVal url_old As String) As String

Dim SysRoot_From_Sys As String '��ȡ�Ļ������� ����ϵͳ�ļ���·��
    If SysRoot <> "" Then
        If SysRoot = "0" Then
            SysRoot_From_Sys = Environ("systemroot")
        Else
            SysRoot_From_Sys = SysRoot
        End If
    Else
        SysRoot_From_Sys = Environ("systemroot")
    End If
    
Dim SysPath_Choose As String
        If SysPath = 0 Then
            SysPath_Choose = "%SystemRoot%"
        ElseIf SysPath = 1 Then
            SysPath_Choose = "%WinDir%"
        ElseIf SysPath = 2 Then
            SysPath_Choose = SysRoot_From_Sys
        Else
            SysPath_Choose = "%SystemRoot%"
        End If

Dim url_now As String
    If InStr(1, url_old, SysRoot_From_Sys, 1) <> 0 Then
        url_now = SysPath_Choose & Mid(url_old, InStr(1, url_old, SysRoot_From_Sys, 1) + Len(SysRoot_From_Sys))
    ElseIf InStr(1, url_old, "%ResourceDir%", 1) <> 0 Then '�ں�ţ�����﷢�ֵġ���
        url_now = SysPath_Choose & "\Resources" & Mid(url_old, InStr(1, url_old, "%ResourceDir%", 1) + Len("%ResourceDir%"))
    Else
        url_now = url_old
    End If
    url_to_P = url_now
End Function

'������������ת����ַ��ͨ��ַ
Public Function url_to_N(ByVal myString As String) As String
Dim SysRoot_From_Sys As String '��ȡ�Ļ������� ����ϵͳ�ļ���·��
    If SysRoot <> "" Then
        If SysRoot = "0" Then
            SysRoot_From_Sys = Environ("systemroot")
        Else
            SysRoot_From_Sys = SysRoot
        End If
    Else
        SysRoot_From_Sys = Environ("systemroot")
    End If
'=====================================================
    
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '����������ʽ
    Set objRegExp = New RegExp
    
    '���ñ��ʽ
    objRegExp.Pattern = "%(\w+)%"
    
    'true���жϴ�Сд
    objRegExp.IgnoreCase = True
    
    'falseֻ������һ��,true����ȫ��
    objRegExp.Global = True
    
    '������Ϊԭ��ַ
    RetStr = myString
    
    '�ȼ���Ƿ��л��������ĵط�
    If (objRegExp.Test(myString) = True) Then
        
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   '' Execute search.
        
        For Each objMatch In colMatches   '' Iterate Matches collection.
            Dim Real_Url_Temp As String
            Dim Path_text As String
            '������ʽ�滻
            Path_text = objRegExp.Replace(objMatch.Value, "$1")
            
'            If IsOK(Path_text, "^HomePath$") Then
'                Real_Url_Temp = Environ("SYSTEMDRIVE") & Environ("HOMEPATH")
'            Else
            If IsOK(Path_text, "^AppTitle$") Then
                Real_Url_Temp = App.Title
            ElseIf IsOK(Path_text, "^AppPath$") Then
                Real_Url_Temp = App.Path
            ElseIf IsOK(Path_text, "^AppEXEName$") Then
                Real_Url_Temp = App.EXEName
            ElseIf IsOK(Path_text, "^ResourceDir$") Then
                Real_Url_Temp = Environ("Windir") & "\Resources"
            ElseIf IsOK(Path_text, "^SystemRoot$") Or IsOK(Path_text, "^WinDir$") Then
                Real_Url_Temp = SysRoot_From_Sys
            Else
                Real_Url_Temp = Environ(Path_text)
            End If
            
            '�ҷ���"\"
            If Right$(Real_Url_Temp, 1) <> "\" Then Real_Url_Temp = Real_Url_Temp & "\"
            '��ԭ����Ļ��������滻Ϊ��ʵ��ַ
            RetStr = Replace(RetStr, objMatch.Value, Real_Url_Temp)
        Next
    End If
    
    '����URL�е�/\
    RetStr = ReplaceText(RetStr, "/", "\") '��/�滻��\
    RetStr = ReplaceText(RetStr, "([^\\/])\\{2,}", "$1\") '�����/���滻��һ��/
    RetStr = ReplaceText(RetStr, "^\\{3,}", "\\") '�����\\��ͷ�ľ�������ַ�滻������\
    
    
    If Right$(RetStr, 2) = ",0" Then
        RetStr = Left$(RetStr, Len(RetStr) - 2)
    End If
    
    url_to_N = RetStr
End Function


'���ַ�����Ϊ������
Public Function text_to_num(ByVal text As String) As String
text_to_num = ReplaceText(text, "\D", "")
'Dim num, Case_temp As String
'Dim i As Variant
'For i = 1 To Len(text)
'    Case_temp = Mid$(text, i, 1)
'    Select Case Case_temp
'        Case 0 To 9
'            num = num & Case_temp
'        Case "."
'            num = num & "."
'        Case "һ"
'            num = num & 1
'        Case "��"
'            num = num & 2
'        Case "��"
'            num = num & 3
'        Case "��"
'            num = num & 4
'        Case "��"
'            num = num & 5
'        Case "��"
'            num = num & 6
'        Case "��"
'            num = num & 7
'        Case "��"
'            num = num & 8
'        Case "��"
'            num = num & 9
'        Case "��", "��"
'            num = num & 0
'        Case Else
'    End Select
'Next
'text_to_num = num
End Function
'�����Ƿ����������������ģ�
Public Function Integer_ok(ByVal num As String) As Boolean
Integer_ok = IsOK(num, "^(-?[1-9]\d{0,}|0)$")
'Dim Case_temp As String
'Dim ok As Boolean
'Dim i As Variant
'ok = True
'If Len(num) < 1 Then
'    ok = False
'End If
'If ok = True Then
'    For i = 1 To Len(num)
'        Case_temp = Mid$(num, i, 1)
'        Select Case Case_temp
'            Case "-"
'                If i <> 1 Or Len(num) < 2 Then
'                    ok = False
'                End If
'            Case 0 To 9
'            Case Else
'                ok = False
'        End Select
'        If ok = False Then Exit For
'    Next
'End If
'If ok = False Then
'    Integer_ok = False
'Else
'    Integer_ok = True
'End If
End Function
'���ַ�����Ϊ��ɫ��
Public Function text_to_color(ByVal text As String) As String
text_to_color = ReplaceText(text, "[^0-9a-fA-Fx]", "")
'Dim color, Case_temp As String
'Dim i As Variant
'For i = 1 To Len(text)
'    Case_temp = Mid$(text, i, 1)
'    Select Case Case_temp
'        Case 0 To 9
'            color = color & Case_temp
'        Case "a" To "f", "A" To "F", "x", "X"
'            color = color & Case_temp
'        Case Else
'    End Select
'Next
'text_to_color = color
End Function
'��ʮ������ת��Ϊ16���Ʋ�����
Public Function x10_to_x16(ByVal num As Long, Optional ByVal Wei As Byte) As String
Dim str_Temp As String
Dim i As Byte
str_Temp = Hex$(num)
    If Wei = Null Then
    ElseIf Len(str_Temp) < Wei Then
        For i = 1 To Wei - Len(str_Temp)
            str_Temp = "0" & str_Temp
        Next
    End If
x10_to_x16 = str_Temp
End Function
'��16������ɫ�ַ������10����
Public Function x16_to_x10(ByVal num As String) As Long

'Dim num_Temp As Long
'Dim Case_temp As String
'Dim i As Variant
'num_Temp = 0
'For i = 1 To Len(num)
'    Case_temp = Mid$(num, i, 1)
'    Select Case Case_temp
'        Case 0 To 9
'            Case_temp = Case_temp
'        Case "a", "A"
'            Case_temp = 10
'        Case "b", "B"
'            Case_temp = 11
'        Case "c", "C"
'            Case_temp = 12
'        Case "d", "D"
'            Case_temp = 13
'        Case "e", "E"
'            Case_temp = 14
'        Case "f", "F"
'            Case_temp = 15
'        Case Else
'            Case_temp = 0
'    End Select
'    num_Temp = num_Temp + Val(Case_temp) * 16 ^ (Len(num) - i)
'Next
If num <> "" Then
    x16_to_x10 = CLng("&H" & num) 'num_Temp
Else
    x16_to_x10 = 0
End If
End Function
'�����Ƿ������ɫ����
Public Function Color_ok(ByVal color As String) As String
Dim Case_temp As String
Dim ok As Boolean
Dim i As Variant
ok = True
If Len(color) = 8 Then
    ok = IsOK(color, "^[0-9a-fA-F]{8}$")
'    For i = 1 To Len(color)
'        Case_temp = Mid$(color, i, 1)
'        Select Case Case_temp
'            Case 0 To 9, "a" To "f", "A" To "F"
'            Case Else
'                ok = False
'        End Select
'        If ok = False Then Exit For
'    Next
ElseIf Len(color) = 10 Then
    ok = IsOK(color, "^0{1}(?:x|X){1}[0-9a-fA-F]{8}$")
'    For i = 1 To Len(color)
'        Case_temp = Mid$(color, i, 1)
'        Select Case Case_temp
'            Case "x", "X"
'                If i <> 2 Then
'                    ok = False
'                End If
'            Case 1 To 9, "a" To "f", "A" To "F"
'                If i = 1 Or i = 2 Then
'                    ok = False
'                End If
'            Case 0
'                If i = 2 Then
'                    ok = False
'                End If
'            Case Else
'                ok = False
'        End Select
'        If ok = False Then Exit For
'    Next
Else
ok = False
End If

If ok = False Then
    Color_ok = 0
Else
    Color_ok = Len(color)
End If
End Function

'��RGBת����BGR
Public Function RGB_To_BGR(ByVal color As Long) As String
Dim url_now As String

    RGB_To_BGR = x16_to_x10(Mid$(x10_to_x16(color, 6), 5, 2) + Mid$(x10_to_x16(color, 6), 3, 2) + Mid$(x10_to_x16(color, 6), 1, 2))

End Function
'��RGBת����BGR2
Public Function RGB_To_BGR_Alpha(ByVal color As Long, ByVal BackColorKey As Long, ByVal Alpha As Byte) As String
Dim url_now As String
Dim BackColor As String
'BackColorKey = GetSysColor(COLOR_BTNFACE) '��ȡϵͳ��ť��ɫ
BackColor = x10_to_x16(BackColorKey, 6)

RGB_To_BGR_Alpha = x16_to_x10( _
                            x10_to_x16(x16_to_x10(Mid$(x10_to_x16(color, 6), 5, 2)) * (Alpha / 255) + x16_to_x10(Mid$(BackColor, 5, 2)) * ((255 - Alpha) / 255), 2) _
                            + x10_to_x16(x16_to_x10(Mid$(x10_to_x16(color, 6), 3, 2)) * (Alpha / 255) + x16_to_x10(Mid$(BackColor, 3, 2)) * ((255 - Alpha) / 255), 2) _
                            + x10_to_x16(x16_to_x10(Mid$(x10_to_x16(color, 6), 1, 2)) * (Alpha / 255) + x16_to_x10(Mid$(BackColor, 1, 2)) * ((255 - Alpha) / 255), 2) _
                            )
End Function
