Attribute VB_Name = "编辑XML"
Option Explicit

' 返回各个节点的值

Public Function GetNodeText(ByVal start_at_node As DOMDocument, ByVal node_name As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '如果没有就使用默认值
        Value_Temp = default_value
    Else
        Value_Temp = value_node(Index).text
        If Not IsMissing(Max) Then  '如果设置了最大数字则当做数字处理
            Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
            
            If Value_Temp > CLng(Max) Then '如果超过最大值则等于最大值，需要先将Max转变回数字才能识别
                Value_Temp = CLng(Max)
            End If

        End If
    End If
    GetNodeText = Value_Temp
End Function

' 返回节点属性
Public Function GetNodeAttribute(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal AttributeName As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Dim value_node_Attribute As IXMLDOMNode
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '如果节点不存在
        Value_Temp = default_value
    Else
        Set value_node_Attribute = value_node(Index).Attributes.getNamedItem(AttributeName)
        If value_node_Attribute Is Nothing Then '如果节点属性不存在
            Value_Temp = default_value
        Else
            Value_Temp = value_node_Attribute.text
            If Not IsMissing(Max) Then  '如果设置了最大数字则当做数字处理
                Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
                
                If Value_Temp > CLng(Max) Then '如果超过最大值则等于最大值，需要先将Max转变回数字才能识别
                    Value_Temp = CLng(Max)
                End If
    
            End If
        End If
    End If
    GetNodeAttribute = Value_Temp
End Function

'从XML读取同名结点个数
Public Function GetAllNode_Lenth(ByVal start_at_node As DOMDocument, ByVal node_name As String) As Long
    Dim i%
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node Is Nothing Then '如果节点不存在
        Value_Temp = 0
    Else
        Value_Temp = value_node.Length
    End If
    GetAllNode_Lenth = Value_Temp
End Function
