Attribute VB_Name = "正则表达式基本函数"
Option Explicit
'存储正则表达式搜索用
Public Type PatternValue
    FirstIndex As Long
    AllValue As String
    InValue() As String
End Type
Public Function SearchText(ByVal text As String, ByVal patrn As String, ByRef Save() As PatternValue, Optional ByVal replStr As String = "$1") As String
'正则表达式搜索
Dim regEx As RegExp
Dim Match As Match
Dim Matchs As MatchCollection
Dim replStrPart() As String

Dim i As Long, j As Long
Set regEx = New RegExp '建立正则表达式
regEx.Pattern = patrn '设置表达式
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Global = True 'false只搜索第一个,true就是全部
Set Matchs = regEx.Execute(text)

replStrPart = Split(replStr, Chr(0)) '使用Chr(0)分开需要搜索的部分的代码
i = 0
For Each Match In Matchs '遍历，并存入实参数组
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

SearchText = Matchs.Count '显示共有几个？
End Function

Public Function ReplaceText(ByVal text As String, ByVal patrn As String, ByVal replStr As String) As String
'正则表达式替换
Dim regEx '建立变量
Set regEx = New RegExp '建立正则表达式
regEx.Pattern = patrn '设置表达式
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Global = True 'false只搜索第一个,true就是全部
ReplaceText = regEx.Replace(text, replStr) '替换命令
End Function

Public Function IsOK(ByVal text As String, ByVal patrn As String) As Boolean
'正则表达式判断是否符合
Dim regEx
IsOK = False
Set regEx = New RegExp
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Pattern = patrn
IsOK = regEx.Test(text)
End Function

'只保留需要的格式，其他的全丢弃
Public Function OnlyRegExp(myString As String, myPattern As String, Optional ByVal default As String = "", Optional ByVal replStr As String = "") As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '建立正则表达式
    Set objRegExp = New RegExp
    
    '设置表达式
    objRegExp.Pattern = myPattern
    
    'true则不判断大小写
    objRegExp.IgnoreCase = True
    
    'false只搜索第一个,true就是全部
    objRegExp.Global = True
    
    '先检查是否有复合的地方
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
