Attribute VB_Name = "������ʽ��������"
Option Explicit
'�洢������ʽ������
Public Type PatternValue
    FirstIndex As Long
    AllValue As String
    InValue() As String
End Type
Public Function SearchText(ByVal text As String, ByVal patrn As String, ByRef Save() As PatternValue, Optional ByVal replStr As String = "$1") As String
'������ʽ����
Dim regEx As RegExp
Dim Match As Match
Dim Matchs As MatchCollection
Dim replStrPart() As String

Dim i As Long, j As Long
Set regEx = New RegExp '����������ʽ
regEx.Pattern = patrn '���ñ��ʽ
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Global = True 'falseֻ������һ��,true����ȫ��
Set Matchs = regEx.Execute(text)

replStrPart = Split(replStr, Chr(0)) 'ʹ��Chr(0)�ֿ���Ҫ�����Ĳ��ֵĴ���
i = 0
For Each Match In Matchs '������������ʵ������
    ReDim Preserve Save(i)
    With Save(i)
        .FirstIndex = Match.FirstIndex
        .AllValue = Match.Value
        ReDim Preserve .InValue(UBound(replStrPart))
        For j = 0 To UBound(replStrPart)
        .InValue(j) = regEx.Replace(Match.Value, replStrPart(j))
        Next
    End With
    i = i + 1
Next

SearchText = Matchs.Count '��ʾ���м�����
End Function

Public Function ReplaceText(ByVal text As String, ByVal patrn As String, ByVal replStr As String) As String
'������ʽ�滻
Dim regEx '��������
Set regEx = New RegExp '����������ʽ
regEx.Pattern = patrn '���ñ��ʽ
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Global = True 'falseֻ������һ��,true����ȫ��
ReplaceText = regEx.Replace(text, replStr) '�滻����
End Function

Public Function IsOK(ByVal text As String, ByVal patrn As String) As Boolean
'������ʽ�ж��Ƿ����
Dim regEx
IsOK = False
Set regEx = New RegExp
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Pattern = patrn
IsOK = regEx.Test(text)
End Function

'ֻ������Ҫ�ĸ�ʽ��������ȫ����
Public Function OnlyRegExp(myString As String, myPattern As String, Optional ByVal default As String = "", Optional ByVal replStr As String = "") As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '����������ʽ
    Set objRegExp = New RegExp
    
    '���ñ��ʽ
    objRegExp.Pattern = myPattern
    
    'true���жϴ�Сд
    objRegExp.IgnoreCase = True
    
    'falseֻ������һ��,true����ȫ��
    objRegExp.Global = True
    
    '�ȼ���Ƿ��и��ϵĵط�
    If (objRegExp.Test(myString) = True) Then

   ''Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.
        
        For Each objMatch In colMatches   ' Iterate Matches collection.
            If replStr = "" Then
                RetStr = RetStr & objMatch.Value
            Else
                RetStr = objRegExp.Replace(objMatch.Value, replStr)
            End If
        Next
    Else
        RetStr = default
    End If
    OnlyRegExp = RetStr
End Function


Public Function Is_File_Directory(ByVal strTmp As String) As Boolean
Is_File_Directory = IsOK(strTmp, FileURL_Parten)
End Function
