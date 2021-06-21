Attribute VB_Name = "DataPagination"
Sub Data_Prev()
'
' 이전 매크로

    '// 이전 데이터 불러오기
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim DataKey As Integer

    Set wb = ThisWorkbook
    Set wsTrade = wb.Sheets("거래명세서")
    
    DataKey = wsTrade.Range("D5") '// 거래명세서번호

    If DataKey = 1 Then
        MsgBox "이전 데이터가 없습니다."
    Else
        wsTrade.Range("D5") = DataKey - 1
    End If
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub Data_Next()
'
' 다음 매크로

    '// 다음 데이터 불러오기
    Dim wb As Workbook
    Dim wsTrade As Worksheet
    Dim wsData As Worksheet
    Dim DataKey As Integer
    Dim LastDataKey As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("데이터")
    Set wsTrade = wb.Sheets("거래명세서")
    
    DataKey = wsTrade.Range("D5") '// 거래명세서번호
    LastDataKey = wsData.Range("a1").CurrentRegion.Rows.Count - 1 '// 거래명세서번호

    If DataKey >= LastDataKey Then
        MsgBox "다음 데이터가 없습니다."
    Else
        wsTrade.Range("D5") = DataKey + 1
    End If

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

