Attribute VB_Name = "DataNew"
Sub Data_New()
'
' 새로작성 매크로

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTrade As Worksheet
    Dim DataKey As Integer
    Dim TradeRows As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("데이터")
    Set wsTrade = wb.Sheets("거래명세서")

    '// 모드
    With wsTrade
        .Range("AE3") = "새로작성"
        .Range("AE4") = "=IFERROR(IF(VLOOKUP(D5, 데이터, 1, FALSE),1,0),0)"
        .Range("AE5") = "=IFERROR(IF(OR(M7=""""),0,1),0)"
        .Range("AE6") = "=IFERROR(IF(OR(C12="""",SUM(M12:P12)=0),0,1),0)"
    End With
    
    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '// 초기화
    With wsTrade
        .Range("M7:N7").ClearContents '// 공급받는자
        .Range("C12:L21").ClearContents '// 거래목록
        .Range("Q12:Q21").ClearContents '// 비고
    End With

    '// 품목
    With wsTrade.Range("C12:E21").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=품목"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// 거래명세서번호
    DataKey = wsData.Range("a1").CurrentRegion.Rows.Count
    With wsTrade
        .Range("D5") = DataKey  '// 거래명세서번호
        .Range("Q5") = "=TODAY()" '// 거래일시
    End With
    
    '// 공급받는자
    With wsTrade
        .Range("M6") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,3,FALSE))" '// 등록번호
        .Range("Q7") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,5,FALSE))" '// 성명
        .Range("M8") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,6,FALSE))" '// 주소
        .Range("M9") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,7,FALSE))" '// 업태
        .Range("Q9") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,8,FALSE))" '// 종목
        .Range("M10") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,11,FALSE))" '// 전화
        .Range("Q10") = "=IF($M$7="""","""",VLOOKUP(M7,거래처,13,FALSE))" '// 팩스
    End With
    
    '// 거래내역
    For K = 1 To 10
        TradeRows = K + 11
        With wsTrade
            .Cells(TradeRows, 2) = "=IF(C" & TradeRows & "="""","""",ROW()-11)"
            .Cells(TradeRows, 13) = "=IFERROR(IF(H" & TradeRows & "*J" & TradeRows & ",H" & TradeRows & "*J" & TradeRows & ",""""),"""")" '// 공급가액
            .Cells(TradeRows, 14) = "=IFERROR(IF(M" & TradeRows & ",M" & TradeRows & "*0.1,""""),"""")" '// 세액
        End With
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


