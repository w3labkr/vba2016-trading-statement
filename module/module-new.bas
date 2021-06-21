Attribute VB_Name = "DataNew"
Sub Data_New()
'
' �����ۼ� ��ũ��

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet
    Dim DataKey As Integer
    Dim TradeRows As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("������")
    Set wsTrade = wb.Sheets("�ŷ�����")

    '// ���
    With wsTrade
        .Range("AE3") = "�����ۼ�"
        .Range("AE4") = "=IFERROR(IF(VLOOKUP(D5, ������, 1, FALSE),1,0),0)"
        .Range("AE5") = "=IFERROR(IF(OR(M7=""""),0,1),0)"
        .Range("AE6") = "=IFERROR(IF(OR(C12="""",SUM(M12:P12)=0),0,1),0)"
    End With
    
    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '// �ʱ�ȭ
    With wsTrade
        .Range("M7:N7").ClearContents '// ���޹޴���
        .Range("C12:L21").ClearContents '// �ŷ����
        .Range("Q12:Q21").ClearContents '// ���
    End With

    '// ǰ��
    With wsTrade.Range("C12:E21").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=ǰ��"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// �ŷ�������ȣ
    DataKey = wsData.Range("a1").CurrentRegion.Rows.Count
    With wsTrade
        .Range("D5") = DataKey  '// �ŷ�������ȣ
        .Range("Q5") = "=TODAY()" '// �ŷ��Ͻ�
    End With
    
    '// ���޹޴���
    With wsTrade
        .Range("M6") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,3,FALSE))" '// ��Ϲ�ȣ
        .Range("Q7") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,5,FALSE))" '// ����
        .Range("M8") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,6,FALSE))" '// �ּ�
        .Range("M9") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,7,FALSE))" '// ����
        .Range("Q9") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,8,FALSE))" '// ����
        .Range("M10") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,11,FALSE))" '// ��ȭ
        .Range("Q10") = "=IF($M$7="""","""",VLOOKUP(M7,�ŷ�ó,13,FALSE))" '// �ѽ�
    End With
    
    '// �ŷ�����
    For K = 1 To 10
        TradeRows = K + 11
        With wsTrade
            .Cells(TradeRows, 2) = "=IF(C" & TradeRows & "="""","""",ROW()-11)"
            .Cells(TradeRows, 13) = "=IFERROR(IF(H" & TradeRows & "*J" & TradeRows & ",H" & TradeRows & "*J" & TradeRows & ",""""),"""")" '// ���ް���
            .Cells(TradeRows, 14) = "=IFERROR(IF(M" & TradeRows & ",M" & TradeRows & "*0.1,""""),"""")" '// ����
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


