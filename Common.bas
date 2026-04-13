Attribute VB_Name = "Common"
Public Function LastRow(ws As Worksheet, Optional indexCol As Long = 1, Optional chaoMode As Boolean = False, Optional searchForBlank As Boolean = False) As Long

    '找最後一列，預設從欄1的資料去決定（可另外指定），資料結構較混亂時可另指定chaoMode，逐列用countA掃描。
    
    With ws
    
        If chaoMode Then
            '混亂模式
            For i = .Rows.Count To 1 Step -1
                If WorksheetFunction.CountA(.Range(.Cells(i, 1), .Cells(i, .Columns.Count))) > 0 Then
                    LastRow = i
                    Exit For
                End If
            Next i
        Else
            '和平模式
            LastRow = .Range(.Cells(.Rows.Count, indexCol), .Cells(.Rows.Count, indexCol)).End(xlUp).Row
        End If
        
    End With
    
    '需要找到沒資料那列嗎？
    If searchForBlank Then
        LastRow = LastRow + 1
    End If
    
    '避免報錯
    If LastRow = 0 Then LastRow = 1

End Function

Public Function LastCol(ws As Worksheet, Optional indexRow As Long = 1, Optional chaoMode As Boolean = False, Optional searchForBlank As Boolean = False) As Long

    '找最後一欄，預設從列1的資料去決定（可另外指定），資料結構較混亂時可另指定chaoMode，逐欄用countA掃描。

    Dim i As Long

    With ws
        If chaoMode Then
            '混亂模式
            For i = .Columns.Count To 1 Step -1
                If WorksheetFunction.CountA(.Range(.Cells(1, i), .Cells(.Rows.Count, i))) > 0 Then
                    LastCol = i
                    Exit For
                End If
            Next i
        Else
            '和平模式
            LastCol = .Range(.Cells(indexRow, .Columns.Count), .Cells(indexRow, .Columns.Count)).End(xlToLeft).Column
        End If
    End With

    '需要找到沒資料的下一欄嗎？
    If searchForBlank Then
        LastCol = LastCol + 1
    End If

    '避免報錯
    If LastCol = 0 Then LastCol = 1

End Function

Public Function NameToRowIndex(nm As String, ws As Worksheet, Optional smart As Boolean = True, Optional startRow As Long = 1, Optional indexCol As Long = 1) As Long

    '在indexCol裡面找列名稱，並給出對應的列號
    '初始化
    NameToRowIndex = -1
    '開找
    With ws
        '先用exact match去找
        On Error Resume Next
        NameToRowIndex = WorksheetFunction.Match(nm, .Range(.Cells(startRow, indexCol), .Cells(.Rows.Count, indexCol)), 0)
        On Error GoTo 0
        '找不到就用like去找
        If NameToRowIndex <= 0 And smart Then
            For i = startRow To .Rows.Count
                If .Cells(i, indexCol).Value Like "*" & nm & "*" Then
                    NameToRowIndex = i
                    Exit For
                End If
            Next i
        End If
    End With
    '錯誤保護
    If NameToRowIndex <= 0 Then NameToRowIndex = -1

End Function

Public Function NameToColIndex(nm As String, ws As Worksheet, Optional smart As Boolean = True, Optional startCol As Long = 1, Optional indexRow As Long = 1) As Long

    '在indexRow裡面找欄名稱，並給出對應的欄號
    '初始化
    '開找
    NameToColIndex = -1
    With ws
        '先用exact match去找
        On Error Resume Next
        NameToColIndex = WorksheetFunction.Match(nm, .Range(.Cells(indexRow, startCol), .Cells(indexRow, .Columns.Count)), 0)
        On Error GoTo 0
        '找不到就用like去找
        If NameToColIndex <= 0 And smart Then
            For i = startCol To .Columns.Count
                If .Cells(indexRow, i).Value Like "*" & nm & "*" Then
                    NameToColIndex = i
                    Exit For
                End If
            Next i
        End If
    End With
    '錯誤保護
    If NameToColIndex <= 0 Then NameToColIndex = -1

End Function
