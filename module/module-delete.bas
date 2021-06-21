Attribute VB_Name = "DataDelete"
Sub Data_Delete()
'
' 삭제하기 매크로

    '// 데이터 삭제
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
    If wsTrade.Range("AE3") <> "불러오기" Then
        MsgBox "불러오기 모드에서만 작동합니다."
        Exit Sub
    ElseIf wsTrade.Range("AE4") = 0 Then
        MsgBox "데이터가 없습니다."
        Exit Sub
    End If
    
    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// 초기화
    wsTrade.Range("M7:N7").ClearContents '// 공급받는자
    wsTrade.Range("C12:Q21").ClearContents '// 거래목록
    
    DataKey = wsTrade.Range("D5") '// 거래명세서번호
    DataRows = DataKey + 1

    With wsData
        '// .Cells(AddRows, 1) = wsTrade.Cells(5, 4) '// 잠금:거래명세서번호
        .Cells(DataRows, 2) = wsTrade.Cells(12, 27) '//참조번호
        .Cells(DataRows, 3) = wsTrade.Cells(5, 17) '// 거래일시
        .Cells(DataRows, 4) = wsTrade.Cells(3, 28) '// 분기
        .Cells(DataRows, 5) = wsTrade.Cells(4, 28) '//년
        .Cells(DataRows, 6) = wsTrade.Cells(5, 28) '// 월
        .Cells(DataRows, 7) = wsTrade.Cells(6, 28) '// 일
        .Cells(DataRows, 8) = wsTrade.Cells(7, 13) '// 상호
        .Cells(DataRows, 9) = wsTrade.Cells(22, 27) '// 참조번호합계
    End With

    MaxItems = 10 '// 참조번호합계
    
    For K = 1 To MaxItems
        DataCols = K * 10
        With wsData
            '// .Cells(DataRows, DataCols) = wsTrade.Cells(11 + K, 27)   '//참조번호K
            .Cells(DataRows, DataCols + 1).ClearContents '//품목K
            .Cells(DataRows, DataCols + 2).ClearContents '//규격K
            .Cells(DataRows, DataCols + 3).ClearContents '//수량K
            .Cells(DataRows, DataCols + 4).ClearContents '//단위K
            .Cells(DataRows, DataCols + 5).ClearContents '//단가K
            .Cells(DataRows, DataCols + 6).ClearContents '//공급가액K
            .Cells(DataRows, DataCols + 7).ClearContents '//세액K
            .Cells(DataRows, DataCols + 8).ClearContents '//합계K
            .Cells(DataRows, DataCols + 9).ClearContents '//비고K
        End With
    Next

    '// Enables screen refreshing.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    '// 상세데이터 삭제
    Data_Delete_Details
    
    '// 삭제 후 필드 초기화
    Data_Load
    
    '// 완료 메시지
    MsgBox "데이터가 삭제 되었습니다."

End Sub

Sub Data_Delete_Details()
'
' 상세데이터 삭제하기 매크로

    '// 상세데이터 삭제
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
    Set wsTrade = wb.Sheets("거래명세서")
    Set wsData = wb.Sheets("데이터")
    Set wsDetails = wb.Sheets("상세데이터")

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    DataKey = wsTrade.Range("D5") '// 거래명세서번호
    DataRKey = Application.VLookup(DataKey, Range("데이터"), 2, False)  '// 데이터 참조번호
    DetailsKey = Application.WorksheetFunction.Match(DataRKey, wsDetails.Range("A:A"), 0)  '// 상세데이터 참조번호
    DataRows = DataKey + 1
    DetailsRows = DetailsKey - 1
    MaxItems = 10 '// 참조번호합계
    
    For K = 1 To MaxItems
        DataCols = K * 10
        With wsDetails
            '// .Cells(DetailsRows + K, 1) = Application.VLookup(DataKey, Range("데이터"), DataCols, False) '// 잠금:참조번호K
            .Cells(DetailsRows + K, 2) = Application.VLookup(DataKey, Range("데이터"), 1, False) '// 거래명세서번호
            .Cells(DetailsRows + K, 3) = Application.VLookup(DataKey, Range("데이터"), 3, False) '// 거래일시
            .Cells(DetailsRows + K, 4) = Application.VLookup(DataKey, Range("데이터"), 4, False) '// 분기
            .Cells(DetailsRows + K, 5) = Application.VLookup(DataKey, Range("데이터"), 5, False) '//년
            .Cells(DetailsRows + K, 6) = Application.VLookup(DataKey, Range("데이터"), 6, False) '// 월
            .Cells(DetailsRows + K, 7) = Application.VLookup(DataKey, Range("데이터"), 7, False) '// 일
            .Cells(DetailsRows + K, 8) = Application.VLookup(DataKey, Range("데이터"), 8, False) '// 상호
            .Cells(DetailsRows + K, 9) = Application.VLookup(DataKey, Range("데이터"), DataCols + 1, False) '//품목K
            .Cells(DetailsRows + K, 10) = Application.VLookup(DataKey, Range("데이터"), DataCols + 2, False) '//규격K
            .Cells(DetailsRows + K, 11) = Application.VLookup(DataKey, Range("데이터"), DataCols + 3, False) '//수량K
            .Cells(DetailsRows + K, 12) = Application.VLookup(DataKey, Range("데이터"), DataCols + 4, False) '//단위K
            .Cells(DetailsRows + K, 13) = Application.VLookup(DataKey, Range("데이터"), DataCols + 5, False) '//단가K
            .Cells(DetailsRows + K, 14) = Application.VLookup(DataKey, Range("데이터"), DataCols + 6, False) '//공급가액K
            .Cells(DetailsRows + K, 15) = Application.VLookup(DataKey, Range("데이터"), DataCols + 7, False) '//세액K
            .Cells(DetailsRows + K, 16) = Application.VLookup(DataKey, Range("데이터"), DataCols + 8, False) '//합계K
            .Cells(DetailsRows + K, 17) = Application.VLookup(DataKey, Range("데이터"), DataCols + 9, False) '//비고K
        End With
    Next
    
    '// Enables screen refreshing.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub






