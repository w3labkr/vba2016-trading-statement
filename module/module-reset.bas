Attribute VB_Name = "DataReset"
Sub Data_Reset()
'
' 데이터 초기화 매크로

    '// 데이터 초기화
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsDetails As Worksheet
    Dim Answer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("데이터")
    Set wsDetails = wb.Sheets("상세데이터")

    '// 확인 메시지
    Answer = MsgBox("정말 데이터를 초기화 하시겠습니까?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If Answer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '//  데이터
    For K = 1 To wsData.Range("A1").CurrentRegion.Rows.Count
        wsData.Rows(K + 1).ClearContents
    Next
    
    '//  상세데이터
    For K = 1 To wsDetails.Range("A1").CurrentRegion.Rows.Count
        wsDetails.Rows(K + 1).ClearContents
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '// 새로작성
    Data_New

    '// 완료 메시지
    MsgBox "데이터가 초기화 되었습니다."
    
End Sub
