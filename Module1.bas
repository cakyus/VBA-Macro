Option Explicit

Sub Generate_SqlWhereIn_FromSelection()

    Dim cel As Range
    Dim selectedRange As Range
    
    Dim celCount As Integer
    Dim sqlWhereIn As String

    Set selectedRange = Application.Selection
    
    celCount = 1
    For Each cel In selectedRange.Cells
        If celCount = 1 Then
            sqlWhereIn = "'" & cel.Value & "'"
        Else
            sqlWhereIn = sqlWhereIn & ",'" & cel.Value & "'"
        End If
        celCount = celCount + 1
    Next cel
    
    Debug.Print sqlWhereIn

End Sub
