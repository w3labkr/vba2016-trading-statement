Attribute VB_Name = "UpdatePivot"
Sub Update_Pivot()
'
' 업데이트 피벗테이블 매크로

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
            
    '// 회사별
    With Worksheets("회사별").PivotTables("회사별")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("거래일시").PivotItems("(blank)").Visible = False
        .PivotFields("규격").PivotItems("(blank)").Visible = False
        .PivotFields("품목").PivotItems("(blank)").Visible = False
        .PivotFields("상호").PivotItems("(blank)").Visible = False
    End With
        
    '// 제품별
    With Worksheets("제품별").PivotTables("제품별")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("거래일시").PivotItems("(blank)").Visible = False
        .PivotFields("규격").PivotItems("(blank)").Visible = False
        .PivotFields("품목").PivotItems("(blank)").Visible = False
        .PivotFields("상호").PivotItems("(blank)").Visible = False
    End With
  
    '// 분기별
    With Worksheets("분기별").PivotTables("분기별")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("분기").PivotItems("(blank)").Visible = False
        .PivotFields("규격").PivotItems("(blank)").Visible = False
        .PivotFields("품목").PivotItems("(blank)").Visible = False
        .PivotFields("상호").PivotItems("(blank)").Visible = False
    End With

    '// 월별
    With Worksheets("월별").PivotTables("월별")
        .PivotCache.Refresh
        .ClearAllFilters
        .PivotFields("월").PivotItems("(blank)").Visible = False
        .PivotFields("규격").PivotItems("(blank)").Visible = False
        .PivotFields("품목").PivotItems("(blank)").Visible = False
        .PivotFields("상호").PivotItems("(blank)").Visible = False
    End With
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '// 완료 메시지
    MsgBox "업데이트가 완료 되었습니다."

End Sub

