Attribute VB_Name = "UpdatePivot"
Sub Update_Pivot()
'
' ������Ʈ �ǹ����̺� ��ũ��

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
            
    '// ȸ�纰
    With Worksheets("ȸ�纰").PivotTables("ȸ�纰")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("�ŷ��Ͻ�").PivotItems("(blank)").Visible = False
        .PivotFields("�԰�").PivotItems("(blank)").Visible = False
        .PivotFields("ǰ��").PivotItems("(blank)").Visible = False
        .PivotFields("��ȣ").PivotItems("(blank)").Visible = False
    End With
        
    '// ��ǰ��
    With Worksheets("��ǰ��").PivotTables("��ǰ��")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("�ŷ��Ͻ�").PivotItems("(blank)").Visible = False
        .PivotFields("�԰�").PivotItems("(blank)").Visible = False
        .PivotFields("ǰ��").PivotItems("(blank)").Visible = False
        .PivotFields("��ȣ").PivotItems("(blank)").Visible = False
    End With
  
    '// �б⺰
    With Worksheets("�б⺰").PivotTables("�б⺰")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("�б�").PivotItems("(blank)").Visible = False
        .PivotFields("�԰�").PivotItems("(blank)").Visible = False
        .PivotFields("ǰ��").PivotItems("(blank)").Visible = False
        .PivotFields("��ȣ").PivotItems("(blank)").Visible = False
    End With

    '// ����
    With Worksheets("����").PivotTables("����")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("��").PivotItems("(blank)").Visible = False
        .PivotFields("�԰�").PivotItems("(blank)").Visible = False
        .PivotFields("ǰ��").PivotItems("(blank)").Visible = False
        .PivotFields("��ȣ").PivotItems("(blank)").Visible = False
    End With
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '// �Ϸ� �޽���
    MsgBox "������Ʈ�� �Ϸ� �Ǿ����ϴ�."

End Sub

