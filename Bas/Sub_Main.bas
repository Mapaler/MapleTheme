Attribute VB_Name = "Sub_Main"
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
(iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
        lngSize As Long
        lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
        On Error Resume Next
        Dim iccex As tagInitCommonControlsEx
        With iccex
           .lngSize = LenB(iccex)
           .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex
        InitCommonControlsVB = (Err.Number = 0)
        On Error GoTo 0
End Function

Sub Main()
    Call InitCommonControlsVB
    frmLoad.Show '注意这时没有加载Form，因此Me是无效的。
End Sub

