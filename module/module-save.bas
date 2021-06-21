Attribute VB_Name = "DataSave"
Sub Data_Save()
'
' 저장하기 매크로

    '// 데이터 저장
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet
    Dim DataKey As Integer
    Dim MaxItems As Integer
    Dim DataRows As Integer
    Dim DataCols As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("데이터")
    Set wsTrade = wb.Sheets("거래명세서")
    
    '// 모드
    If wsTrade.Range("AE3") <> "새로작성" Then
        MsgBox "새로작성 모드에서만 작동합니다."
        Exit Sub
    ElseIf wsTrade.Range("AE5") = 0 Then
        MsgBox "상호를 입력해 주세요."
        Exit Sub
    ElseIf wsTrade.Range("AE6") = 0 Then
        MsgBox "한개 이상의 제품을 입력해 주세요."
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
        
    DataKey = wsTrade.Range("D5") '// 거래명세서번호
    DataRows = DataKey + 1
    MaxItems = wsTrade.Cells(22, 27) '// 참조번호합계

    With wsData
        .Cells(DataRows, 1) = wsTrade.Cells(5, 4) '// 거래명세서번호
        .Cells(DataRows, 2) = wsTrade.Cells(12, 27) '//참조번호
        .Cells(DataRows, 3) = wsTrade.Cells(5, 17) '// 거래일시
        .Cells(DataRows, 4) = wsTrade.Cells(3, 28) '// 분기
        .Cells(DataRows, 5) = wsTrade.Cells(4, 28) '//년
        .Cells(DataRows, 6) = wsTrade.Cells(5, 28) '// 월
        .Cells(DataRows, 7) = wsTrade.Cells(6, 28) '// 일
        .Cells(DataRows, 8) = wsTrade.Cells(7, 13) '// 상호
        .Cells(DataRows, 9) = wsTrade.Cells(22, 27) '// 참조번호합계
    End With

    For K = 1 To MaxItems
        DataCols = K * 10
        With wsData
            .Cells(DataRows, DataCols) = wsTrade.Cells(11 + K, 27) '//참조번호K
            .Cells(DataRows, DataCols + 1) = wsTrade.Cells(11 + K, 3) '//품목K
            .Cells(DataRows, DataCols + 2) = wsTrade.Cells(11 + K, 6) '//규격K
            .Cells(DataRows, DataCols + 3) = wsTrade.Cells(11 + K, 8) '//수량K
            .Cells(DataRows, DataCols + 4) = wsTrade.Cells(11 + K, 9) '//단위K
            .Cells(DataRows, DataCols + 5) = wsTrade.Cells(11 + K, 10) '//단가K
            .Cells(DataRows, DataCols + 6) = wsTrade.Cells(11 + K, 13) '//공급가액K
            .Cells(DataRows, DataCols + 7) = wsTrade.Cells(11 + K, 14) '//세액K
            .Cells(DataRows, DataCols + 8) = wsTrade.Cells(11 + K, 28) '//합계K
            .Cells(DataRows, DataCols + 9) = wsTrade.Cells(11 + K, 17) '//비고K
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '// 상세데이터 저장
    Data_Save_Details

    '// 데이터 불러오기
    Data_Load

    '// 완료 메시지
    MsgBox "데이터가 저장 되었습니다."

End Sub

Sub Data_Save_Details()
'
' 상세데이터 저장하기 매크로

    '// 상세데이터
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
    Set wsTrade = wb.Sheets("거래명세서")
    Set wsData = wb.Sheets("데이터")
    Set wsDetails = wb.Sheets("상세데이터")

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    DataKey = wsTrade.Range("D5") '// 거래명세서번호
    DetailsKey = wsDetails.Range("A1").CurrentRegion.Rows.Count  '//상세데이터 일련번호
    DataRows = DataKey + 1
    DetailsRows = DetailsKey
    MaxItems = wsData.Cells(DataRows, 9)    '// 참조번호합계
    
    For K = 1 To MaxItems
        EditCols = K * 10
        With wsDetails
            .Cells(DetailsRows + K, 1) = wsData.Cells(DataRows, EditCols) '//참조번호K
            .Cells(DetailsRows + K, 2) = wsData.Cells(DataRows, 1) '// 거래명세서번호
            .Cells(DetailsRows + K, 3) = wsData.Cells(DataRows, 3) '// 거래일시
            .Cells(DetailsRows + K, 4) = wsData.Cells(DataRows, 4) '// 분기
            .Cells(DetailsRows + K, 5) = wsData.Cells(DataRows, 5) '//년
            .Cells(DetailsRows + K, 6) = wsData.Cells(DataRows, 6) '// 월
            .Cells(DetailsRows + K, 7) = wsData.Cells(DataRows, 7) '// 일
            .Cells(DetailsRows + K, 8) = wsData.Cells(DataRows, 8) '// 상호
            .Cells(DetailsRows + K, 9) = wsData.Cells(DataRows, EditCols + 1) '//품목K
            .Cells(DetailsRows + K, 10) = wsData.Cells(DataRows, EditCols + 2) '//규격K
            .Cells(DetailsRows + K, 11) = wsData.Cells(DataRows, EditCols + 3) '//수량K
            .Cells(DetailsRows + K, 12) = wsData.Cells(DataRows, EditCols + 4) '//단위K
            .Cells(DetailsRows + K, 13) = wsData.Cells(DataRows, EditCols + 5) '//단가K
            .Cells(DetailsRows + K, 14) = wsData.Cells(DataRows, EditCols + 6) '//공급가액K
            .Cells(DetailsRows + K, 15) = wsData.Cells(DataRows, EditCols + 7) '//세액K
            .Cells(DetailsRows + K, 16) = wsData.Cells(DataRows, EditCols + 8) '//합계K
            .Cells(DetailsRows + K, 17) = wsData.Cells(DataRows, EditCols + 9) '//비고K
        End With
    Next

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


