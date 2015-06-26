Attribute VB_Name = "�༭XML"
Option Explicit

' ���ظ����ڵ��ֵ

Public Function GetNodeText(ByVal start_at_node As DOMDocument, ByVal node_name As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '���û�о�ʹ��Ĭ��ֵ
        Value_Temp = default_value
    Else
        Value_Temp = value_node(Index).text
        If Not IsMissing(Max) Then  '�����������������������ִ���
            Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
            
            If Value_Temp > CLng(Max) Then '����������ֵ��������ֵ����Ҫ�Ƚ�Maxת������ֲ���ʶ��
                Value_Temp = CLng(Max)
            End If

        End If
    End If
    GetNodeText = Value_Temp
End Function

' ���ؽڵ�����
Public Function GetNodeAttribute(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal AttributeName As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Dim value_node_Attribute As IXMLDOMNode
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '����ڵ㲻����
        Value_Temp = default_value
    Else
        Set value_node_Attribute = value_node(Index).Attributes.getNamedItem(AttributeName)
        If value_node_Attribute Is Nothing Then '����ڵ����Բ�����
            Value_Temp = default_value
        Else
            Value_Temp = value_node_Attribute.text
            If Not IsMissing(Max) Then  '�����������������������ִ���
                Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
                
                If Value_Temp > CLng(Max) Then '����������ֵ��������ֵ����Ҫ�Ƚ�Maxת������ֲ���ʶ��
                    Value_Temp = CLng(Max)
                End If
    
            End If
        End If
    End If
    GetNodeAttribute = Value_Temp
End Function

'��XML��ȡͬ��������
Public Function GetAllNode_Lenth(ByVal start_at_node As DOMDocument, ByVal node_name As String) As Long
    Dim i%
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node Is Nothing Then '����ڵ㲻����
        Value_Temp = 0
    Else
        Value_Temp = value_node.Length
    End If
    GetAllNode_Lenth = Value_Temp
End Function
