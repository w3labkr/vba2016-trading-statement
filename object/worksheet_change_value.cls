VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 직접입력 매크로
'
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Skip
    
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("C12:G21")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        ' Display a message when one of the designated cells has been
        ' changed.
        ' Place your code here.
        ' MsgBox "Cell " & Target.Address & " has changed."
       
        If Target.Value = "직접입력" Then    '// 셀값에 "직접입력" 이 입력되면~~
           Target.Validation.Delete              '// 유효성검사를 하지 않도록 한다
        End If
    End If
    
Done:
    Exit Sub
Skip:
End Sub



