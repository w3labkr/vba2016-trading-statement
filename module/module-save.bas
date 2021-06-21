Attribute VB_Name = "DataSave"
Sub Data_Save()
'
' �����ϱ� ��ũ��

    '// ������ ����
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet
    Dim DataKey As Integer
    Dim MaxItems As Integer
    Dim DataRows As Integer
    Dim DataCols As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("������")
    Set wsTrade = wb.Sheets("�ŷ�����")
    
    '// ���
    If wsTrade.Range("AE3") <> "�����ۼ�" Then
        MsgBox "�����ۼ� ��忡���� �۵��մϴ�."
        Exit Sub
    ElseIf wsTrade.Range("AE5") = 0 Then
        MsgBox "��ȣ�� �Է��� �ּ���."
        Exit Sub
    ElseIf wsTrade.Range("AE6") = 0 Then
        MsgBox "�Ѱ� �̻��� ��ǰ�� �Է��� �ּ���."
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
        
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ
    DataRows = DataKey + 1
    MaxItems = wsTrade.Cells(22, 27) '// ������ȣ�հ�

    With wsData
        .Cells(DataRows, 1) = wsTrade.Cells(5, 4) '// �ŷ�������ȣ
        .Cells(DataRows, 2) = wsTrade.Cells(12, 27) '//������ȣ
        .Cells(DataRows, 3) = wsTrade.Cells(5, 17) '// �ŷ��Ͻ�
        .Cells(DataRows, 4) = wsTrade.Cells(3, 28) '// �б�
        .Cells(DataRows, 5) = wsTrade.Cells(4, 28) '//��
        .Cells(DataRows, 6) = wsTrade.Cells(5, 28) '// ��
        .Cells(DataRows, 7) = wsTrade.Cells(6, 28) '// ��
        .Cells(DataRows, 8) = wsTrade.Cells(7, 13) '// ��ȣ
        .Cells(DataRows, 9) = wsTrade.Cells(22, 27) '// ������ȣ�հ�
    End With

    For K = 1 To MaxItems
        DataCols = K * 10
        With wsData
            .Cells(DataRows, DataCols) = wsTrade.Cells(11 + K, 27) '//������ȣK
            .Cells(DataRows, DataCols + 1) = wsTrade.Cells(11 + K, 3) '//ǰ��K
            .Cells(DataRows, DataCols + 2) = wsTrade.Cells(11 + K, 6) '//�԰�K
            .Cells(DataRows, DataCols + 3) = wsTrade.Cells(11 + K, 8) '//����K
            .Cells(DataRows, DataCols + 4) = wsTrade.Cells(11 + K, 9) '//����K
            .Cells(DataRows, DataCols + 5) = wsTrade.Cells(11 + K, 10) '//�ܰ�K
            .Cells(DataRows, DataCols + 6) = wsTrade.Cells(11 + K, 13) '//���ް���K
            .Cells(DataRows, DataCols + 7) = wsTrade.Cells(11 + K, 14) '//����K
            .Cells(DataRows, DataCols + 8) = wsTrade.Cells(11 + K, 28) '//�հ�K
            .Cells(DataRows, DataCols + 9) = wsTrade.Cells(11 + K, 17) '//���K
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '// �󼼵����� ����
    Data_Save_Details

    '// ������ �ҷ�����
    Data_Load

    '// �Ϸ� �޽���
    MsgBox "�����Ͱ� ���� �Ǿ����ϴ�."

End Sub

Sub Data_Save_Details()
'
' �󼼵����� �����ϱ� ��ũ��

    '// �󼼵�����
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim wsData As Worksheet
    Dim wsDetails As Worksheet
    Dim DataKey As Integer
    Dim DetailsKey As Integer
    Dim DataRows As Integer
    Dim DetailsRows As Integer
    Dim EditCols As Integer
    Dim MaxItems As Integer

    Set wb = ThisWorkbook
    Set wsTrade = wb.Sheets("�ŷ�����")
    Set wsData = wb.Sheets("������")
    Set wsDetails = wb.Sheets("�󼼵�����")

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ
    DetailsKey = wsDetails.Range("A1").CurrentRegion.Rows.Count  '//�󼼵����� �Ϸù�ȣ
    DataRows = DataKey + 1
    DetailsRows = DetailsKey
    MaxItems = wsData.Cells(DataRows, 9)    '// ������ȣ�հ�
    
    For K = 1 To MaxItems
        EditCols = K * 10
        With wsDetails
            .Cells(DetailsRows + K, 1) = wsData.Cells(DataRows, EditCols) '//������ȣK
            .Cells(DetailsRows + K, 2) = wsData.Cells(DataRows, 1) '// �ŷ�������ȣ
            .Cells(DetailsRows + K, 3) = wsData.Cells(DataRows, 3) '// �ŷ��Ͻ�
            .Cells(DetailsRows + K, 4) = wsData.Cells(DataRows, 4) '// �б�
            .Cells(DetailsRows + K, 5) = wsData.Cells(DataRows, 5) '//��
            .Cells(DetailsRows + K, 6) = wsData.Cells(DataRows, 6) '// ��
            .Cells(DetailsRows + K, 7) = wsData.Cells(DataRows, 7) '// ��
            .Cells(DetailsRows + K, 8) = wsData.Cells(DataRows, 8) '// ��ȣ
            .Cells(DetailsRows + K, 9) = wsData.Cells(DataRows, EditCols + 1) '//ǰ��K
            .Cells(DetailsRows + K, 10) = wsData.Cells(DataRows, EditCols + 2) '//�԰�K
            .Cells(DetailsRows + K, 11) = wsData.Cells(DataRows, EditCols + 3) '//����K
            .Cells(DetailsRows + K, 12) = wsData.Cells(DataRows, EditCols + 4) '//����K
            .Cells(DetailsRows + K, 13) = wsData.Cells(DataRows, EditCols + 5) '//�ܰ�K
            .Cells(DetailsRows + K, 14) = wsData.Cells(DataRows, EditCols + 6) '//���ް���K
            .Cells(DetailsRows + K, 15) = wsData.Cells(DataRows, EditCols + 7) '//����K
            .Cells(DetailsRows + K, 16) = wsData.Cells(DataRows, EditCols + 8) '//�հ�K
            .Cells(DetailsRows + K, 17) = wsData.Cells(DataRows, EditCols + 9) '//���K
        End With
    Next

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


