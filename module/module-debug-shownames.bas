Attribute VB_Name = "DebugShowNames"
Sub Debug_Show_Names()
'/// (������) ��� �̸��� ���̰� ��
Dim n As Name
For Each n In ThisWorkbook.Names
    n.Visible = True
Next n
End Sub
