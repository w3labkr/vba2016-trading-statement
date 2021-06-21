Attribute VB_Name = "DataLoad"
Sub Data_Load()
'
' �ҷ����� ��ũ��

    '// ������ �ҷ�����
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("������")
    Set wsTrade = wb.Sheets("�ŷ�����")
    
    '// ���
    wsTrade.Range("AE3") = "�ҷ�����"

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// �ʱ�ȭ
    With wsTrade
        .Range("M6:Q6").ClearContents '// ��Ϲ�ȣ
        .Range("M7:N7").ClearContents '// ��ȣ
        .Range("M8:Q8").ClearContents '// �ּ�
        .Range("M9:N9").ClearContents '// ����
        .Range("M10:N10").ClearContents '// ��ȭ
        .Range("Q5").ClearContents '// �ŷ��Ͻ�
        .Range("Q7").ClearContents '// ����
        .Range("Q9").ClearContents '// ����
        .Range("Q10").ClearContents '// �ѽ�
        .Range("C12:Q21").ClearContents '// �ŷ�����
    End With

    '// ���޹޴���
    With wsTrade
        .Range("Q5") = "=IF(VLOOKUP(D5,������,2,FALSE)="""","""",VLOOKUP(D5,������,3,FALSE))" '// �ŷ��Ͻ�
        .Range("M7:N7") = "=IF(VLOOKUP(D5,������,8,FALSE)="""","""",VLOOKUP(D5,������,8,FALSE))" '// ��ȣ
        .Range("M6") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,3,FALSE))" '// ��Ϲ�ȣ
        .Range("Q7") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,5,FALSE))" '// ����
        .Range("M8") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,6,FALSE))" '// �ּ�
        .Range("M9") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,7,FALSE))" '// ����
        .Range("Q9") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,8,FALSE))" '// ����
        .Range("M10") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,11,FALSE))" '// ��ȭ
        .Range("Q10") = "=IF($M$7="""","""",VLOOKUP($M$7,�ŷ�ó,13,FALSE))" '// �ѽ�
    End With
    
    '// �ŷ�����
    Dim TradeRows As Integer
    Dim DataCols As Integer
    Dim MaxItem As Integer
    
    MaxItem = 10 '// ������ȣ�հ�
    
    For K = 1 To MaxItem
        TradeRows = K + 11
        DataCols = K * 10
        With wsTrade
            .Cells(TradeRows, 3) = "=IF(VLOOKUP(D5,������," & DataCols + 1 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 1 & ",FALSE))" '//ǰ��
            .Cells(TradeRows, 6) = "=IF(VLOOKUP(D5,������," & DataCols + 2 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 2 & ",FALSE))" '//�԰�
            .Cells(TradeRows, 8) = "=IF(VLOOKUP(D5,������," & DataCols + 3 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 3 & ",FALSE))" '//����
            .Cells(TradeRows, 9) = "=IF(VLOOKUP(D5,������," & DataCols + 4 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 4 & ",FALSE))" '//����
            .Cells(TradeRows, 10) = "=IF(VLOOKUP(D5,������," & DataCols + 5 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 5 & ",FALSE))" '//�ܰ�
            .Cells(TradeRows, 17) = "=IF(VLOOKUP(D5,������," & DataCols + 9 & ",FALSE)="""","""",VLOOKUP(D5,������," & DataCols + 9 & ",FALSE))" '//���
            .Cells(TradeRows, 13) = "=IFERROR(IF(H" & TradeRows & "*J" & TradeRows & ",H" & TradeRows & "*J" & TradeRows & ",""""),"""")" '// ���ް���
            .Cells(TradeRows, 14) = "=IFERROR(IF(M" & TradeRows & ",M" & TradeRows & "*0.1,""""),"""")" '// ����
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

