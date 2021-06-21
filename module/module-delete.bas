Attribute VB_Name = "DataDelete"
Sub Data_Delete()
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
    If wsTrade.Range("AE3") <> "�ҷ�����" Then
        MsgBox "�ҷ����� ��忡���� �۵��մϴ�."
        Exit Sub
    ElseIf wsTrade.Range("AE4") = 0 Then
        MsgBox "�����Ͱ� �����ϴ�."
        Exit Sub
    End If
    
    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// �ʱ�ȭ
    wsTrade.Range("M7:N7").ClearContents '// ���޹޴���
    wsTrade.Range("C12:Q21").ClearContents '// �ŷ����
    
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ
    DataRows = DataKey + 1

    With wsData
        '// .Cells(AddRows, 1) = wsTrade.Cells(5, 4) '// ���:�ŷ�������ȣ
        .Cells(DataRows, 2) = wsTrade.Cells(12, 27) '//������ȣ
        .Cells(DataRows, 3) = wsTrade.Cells(5, 17) '// �ŷ��Ͻ�
        .Cells(DataRows, 4) = wsTrade.Cells(3, 28) '// �б�
        .Cells(DataRows, 5) = wsTrade.Cells(4, 28) '//��
        .Cells(DataRows, 6) = wsTrade.Cells(5, 28) '// ��
        .Cells(DataRows, 7) = wsTrade.Cells(6, 28) '// ��
        .Cells(DataRows, 8) = wsTrade.Cells(7, 13) '// ��ȣ
        .Cells(DataRows, 9) = wsTrade.Cells(22, 27) '// ������ȣ�հ�
    End With

    MaxItems = 10 '// ������ȣ�հ�
    
    For K = 1 To MaxItems
        DataCols = K * 10
        With wsData
            '// .Cells(DataRows, DataCols) = wsTrade.Cells(11 + K, 27)   '//������ȣK
            .Cells(DataRows, DataCols + 1).ClearContents '//ǰ��K
            .Cells(DataRows, DataCols + 2).ClearContents '//�԰�K
            .Cells(DataRows, DataCols + 3).ClearContents '//����K
            .Cells(DataRows, DataCols + 4).ClearContents '//����K
            .Cells(DataRows, DataCols + 5).ClearContents '//�ܰ�K
            .Cells(DataRows, DataCols + 6).ClearContents '//���ް���K
            .Cells(DataRows, DataCols + 7).ClearContents '//����K
            .Cells(DataRows, DataCols + 8).ClearContents '//�հ�K
            .Cells(DataRows, DataCols + 9).ClearContents '//���K
        End With
    Next

    '// Enables screen refreshing.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    '// �󼼵����� ����
    Data_Delete_Details
    
    '// ���� �� �ʵ� �ʱ�ȭ
    Data_Load
    
    '// �Ϸ� �޽���
    MsgBox "�����Ͱ� ���� �Ǿ����ϴ�."

End Sub

Sub Data_Delete_Details()
'
' �󼼵����� �����ϱ� ��ũ��

    '// �󼼵����� ����
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim wsData As Worksheet
    Dim wsDetails As Worksheet
    Dim DataKey As Integer
    Dim DataRKey As Integer
    Dim DataRows As Integer
    Dim DataCols As Integer
    Dim DetailsKey As Integer
    Dim DetailsRows As Integer
    Dim MaxItems As Integer

    Set wb = ThisWorkbook
    Set wsTrade = wb.Sheets("�ŷ�����")
    Set wsData = wb.Sheets("������")
    Set wsDetails = wb.Sheets("�󼼵�����")

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ
    DataRKey = Application.VLookup(DataKey, Range("������"), 2, False)  '// ������ ������ȣ
    DetailsKey = Application.WorksheetFunction.Match(DataRKey, wsDetails.Range("A:A"), 0)  '// �󼼵����� ������ȣ
    DataRows = DataKey + 1
    DetailsRows = DetailsKey - 1
    MaxItems = 10 '// ������ȣ�հ�
    
    For K = 1 To MaxItems
        DataCols = K * 10
        With wsDetails
            '// .Cells(DetailsRows + K, 1) = Application.VLookup(DataKey, Range("������"), DataCols, False) '// ���:������ȣK
            .Cells(DetailsRows + K, 2) = Application.VLookup(DataKey, Range("������"), 1, False) '// �ŷ�������ȣ
            .Cells(DetailsRows + K, 3) = Application.VLookup(DataKey, Range("������"), 3, False) '// �ŷ��Ͻ�
            .Cells(DetailsRows + K, 4) = Application.VLookup(DataKey, Range("������"), 4, False) '// �б�
            .Cells(DetailsRows + K, 5) = Application.VLookup(DataKey, Range("������"), 5, False) '//��
            .Cells(DetailsRows + K, 6) = Application.VLookup(DataKey, Range("������"), 6, False) '// ��
            .Cells(DetailsRows + K, 7) = Application.VLookup(DataKey, Range("������"), 7, False) '// ��
            .Cells(DetailsRows + K, 8) = Application.VLookup(DataKey, Range("������"), 8, False) '// ��ȣ
            .Cells(DetailsRows + K, 9) = Application.VLookup(DataKey, Range("������"), DataCols + 1, False) '//ǰ��K
            .Cells(DetailsRows + K, 10) = Application.VLookup(DataKey, Range("������"), DataCols + 2, False) '//�԰�K
            .Cells(DetailsRows + K, 11) = Application.VLookup(DataKey, Range("������"), DataCols + 3, False) '//����K
            .Cells(DetailsRows + K, 12) = Application.VLookup(DataKey, Range("������"), DataCols + 4, False) '//����K
            .Cells(DetailsRows + K, 13) = Application.VLookup(DataKey, Range("������"), DataCols + 5, False) '//�ܰ�K
            .Cells(DetailsRows + K, 14) = Application.VLookup(DataKey, Range("������"), DataCols + 6, False) '//���ް���K
            .Cells(DetailsRows + K, 15) = Application.VLookup(DataKey, Range("������"), DataCols + 7, False) '//����K
            .Cells(DetailsRows + K, 16) = Application.VLookup(DataKey, Range("������"), DataCols + 8, False) '//�հ�K
            .Cells(DetailsRows + K, 17) = Application.VLookup(DataKey, Range("������"), DataCols + 9, False) '//���K
        End With
    Next
    
    '// Enables screen refreshing.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub






