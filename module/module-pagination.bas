Attribute VB_Name = "DataPagination"
Sub Data_Prev()
'
' ���� ��ũ��

    '// ���� ������ �ҷ�����
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim DataKey As Integer

    Set wb = ThisWorkbook
    Set wsTrade = wb.Sheets("�ŷ�����")
    
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ

    If DataKey = 1 Then
        MsgBox "���� �����Ͱ� �����ϴ�."
    Else
        wsTrade.Range("D5") = DataKey - 1
    End If
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub Data_Next()
'
' ���� ��ũ��

    '// ���� ������ �ҷ�����
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim wsData As Worksheet
    Dim DataKey As Integer
    Dim LastDataKey As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("������")
    Set wsTrade = wb.Sheets("�ŷ�����")
    
    DataKey = wsTrade.Range("D5") '// �ŷ�������ȣ
    LastDataKey = wsData.Range("a1").CurrentRegion.Rows.Count - 1 '// �ŷ�������ȣ

    If DataKey >= LastDataKey Then
        MsgBox "���� �����Ͱ� �����ϴ�."
    Else
        wsTrade.Range("D5") = DataKey + 1
    End If

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

