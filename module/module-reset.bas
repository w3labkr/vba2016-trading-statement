Attribute VB_Name = "DataReset"
Sub Data_Reset()
'
' ������ �ʱ�ȭ ��ũ��

    '// ������ �ʱ�ȭ
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsDetails As Worksheet
    Dim Answer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("������")
    Set wsDetails = wb.Sheets("�󼼵�����")

    '// Ȯ�� �޽���
    Answer = MsgBox("���� �����͸� �ʱ�ȭ �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If Answer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '//  ������
    For K = 1 To wsData.Range("A1").CurrentRegion.Rows.Count
        wsData.Rows(K + 1).ClearContents
    Next
    
    '//  �󼼵�����
    For K = 1 To wsDetails.Range("A1").CurrentRegion.Rows.Count
        wsDetails.Rows(K + 1).ClearContents
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '// �����ۼ�
    Data_New

    '// �Ϸ� �޽���
    MsgBox "�����Ͱ� �ʱ�ȭ �Ǿ����ϴ�."
    
End Sub
