Attribute VB_Name = "DataLoad"
Sub Data_Load()
'
' 불러오기 매크로

    '// 데이터 불러오기
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("데이터")
    Set wsTrade = wb.Sheets("거래명세서")
    
    '// 모드
    wsTrade.Range("AE3") = "불러오기"

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// 초기화
    With wsTrade
        .Range("M6:Q6").ClearContents '// 등록번호
        .Range("M7:N7").ClearContents '// 상호
        .Range("M8:Q8").ClearContents '// 주소
        .Range("M9:N9").ClearContents '// 업태
        .Range("M10:N10").ClearContents '// 전화
        .Range("Q5").ClearContents '// 거래일시
        .Range("Q7").ClearContents '// 성명
        .Range("Q9").ClearContents '// 종목
        .Range("Q10").ClearContents '// 팩스
        .Range("C12:Q21").ClearContents '// 거래내역
    End With

    '// 공급받는자
    With wsTrade
        .Range("Q5") = "=IF(VLOOKUP(D5,데이터,2,FALSE)="""","""",VLOOKUP(D5,데이터,3,FALSE))" '// 거래일시
        .Range("M7:N7") = "=IF(VLOOKUP(D5,데이터,8,FALSE)="""","""",VLOOKUP(D5,데이터,8,FALSE))" '// 상호
        .Range("M6") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,3,FALSE))" '// 등록번호
        .Range("Q7") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,5,FALSE))" '// 성명
        .Range("M8") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,6,FALSE))" '// 주소
        .Range("M9") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,7,FALSE))" '// 업태
        .Range("Q9") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,8,FALSE))" '// 종목
        .Range("M10") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,11,FALSE))" '// 전화
        .Range("Q10") = "=IF($M$7="""","""",VLOOKUP($M$7,거래처,13,FALSE))" '// 팩스
    End With
    
    '// 거래내역
    Dim TradeRows As Integer
    Dim DataCols As Integer
    Dim MaxItem As Integer
    
    MaxItem = 10 '// 참조번호합계
    
    For K = 1 To MaxItem
        TradeRows = K + 11
        DataCols = K * 10
        With wsTrade
            .Cells(TradeRows, 3) = "=IF(VLOOKUP(D5,데이터," & DataCols + 1 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 1 & ",FALSE))" '//품목
            .Cells(TradeRows, 6) = "=IF(VLOOKUP(D5,데이터," & DataCols + 2 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 2 & ",FALSE))" '//규격
            .Cells(TradeRows, 8) = "=IF(VLOOKUP(D5,데이터," & DataCols + 3 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 3 & ",FALSE))" '//수량
            .Cells(TradeRows, 9) = "=IF(VLOOKUP(D5,데이터," & DataCols + 4 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 4 & ",FALSE))" '//단위
            .Cells(TradeRows, 10) = "=IF(VLOOKUP(D5,데이터," & DataCols + 5 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 5 & ",FALSE))" '//단가
            .Cells(TradeRows, 17) = "=IF(VLOOKUP(D5,데이터," & DataCols + 9 & ",FALSE)="""","""",VLOOKUP(D5,데이터," & DataCols + 9 & ",FALSE))" '//비고
            .Cells(TradeRows, 13) = "=IFERROR(IF(H" & TradeRows & "*J" & TradeRows & ",H" & TradeRows & "*J" & TradeRows & ",""""),"""")" '// 공급가액
            .Cells(TradeRows, 14) = "=IFERROR(IF(M" & TradeRows & ",M" & TradeRows & "*0.1,""""),"""")" '// 세액
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

