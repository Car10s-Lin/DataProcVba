Attribute VB_Name = "DataFilter"
Public Sub BulkAutoFilter(ws As Worksheet, filterCol As Variant, crit1 As Variant, crit2 As Variant, op As Variant, Optional smartColNum As Boolean = True, Optional startRow As Long = 1, Optional startCol As Long = 1, Optional endRow As Long = 1048576)
    
    '批次Autofilter篩選，傳入條件陣列
    'smartColNum預設true，filterCol可傳入欄位名稱，以第一列為標題列找欄號
    '亦可直接傳入欄號，smartColNum = false
    '可指定資料範圍（預設全表）
    
    '重置filter
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
    
    '檢查各項條件的array是否數量相等
    If Not UBound(filterCol) = UBound(crit1) And UBound(filterCol) = UBound(crit2) And UBound(filterCol) = UBound(op) Then
        MsgBox "喵，篩選條件數目不相等喔，再檢查一下"
        Exit Sub
    End If
    
    'smartColNum置換
    If smartColNum Then
        tempFilterCol = filterCol
        ReDim filterCol(LBound(tempFilterCol) To UBound(tempFilterCol))
        With ws
            For i = LBound(tempFilterCol) To UBound(tempFilterCol)
                filterCol(i) = Application.Match(tempFilterCol(i), .Range(.Cells(startRow, startCol), .Cells(startRow, .Columns.Count)), 0)
            Next i
        End With
    End If
    
    '篩選
    With ws
        With .Range(.Cells(startRow, startCol), .Cells(endRow, 16384))
            For i = LBound(filterCol) To UBound(filterCol)
                If crit2(i) = "" And op(i) = "" Then
                    .AutoFilter field:=filterCol(i), Criteria1:=crit1(i)
                ElseIf crit2(i) = "" Then
                    .AutoFilter field:=filterCol(i), Criteria1:=crit1(i), Operator:=op(i)
                ElseIf op(i) = "" Then
                    .AutoFilter field:=filterCol(i), Criteria1:=crit1(i), Criteria2:=crit2(i)
                Else
                    .AutoFilter field:=filterCol(i), Criteria1:=crit1(i), Criteria2:=crit2(i), Operator:=op(i)
                End If
            Next i
        End With
    End With
    
End Sub

Public Function NthQuantile(ws As Worksheet, portion As Long, refCol As Long, Optional outputCol As Long = -1, Optional titleRows As Long = 1, Optional outputRank As Boolean = False) As Variant

    '根據refCol的排名產生portion分位數陣列
    '可設定計算時扣除titleRows列的資料（標題列／空白列）
    '可選擇輸出分位數連結的任意欄位，預設為排序欄本身
    '亦可選擇僅輸出排名（outputRank = True）
    
    '初始化outputCol
    If outputCol < 1 Then outputCol = refCol
    '先算n分位數應有的名次
    Dim quantiles As Variant
    quantiles = Array()
    For i = 0 To portion - 2
        ReDim Preserve quantiles(i)
        quantiles(i) = CInt((LastRow(ws, refCol) - titleRows) / 5) * (i + 1)
        Debug.Print quantiles(i)
    Next i
    '獲取n分位排名來源的array
    Dim quantileRefs As Variant
    quantileRefs = Array()
    ReDim quantileRefs(UBound(quantiles))
    For i = 2 To LastRow(ws, refCol)
        For j = LBound(quantiles) To UBound(quantiles)
            If WorksheetFunction.Rank(CLng(ws.Cells(i, refCol).Value), ws.Range(ws.Cells(i, refCol), ws.Cells(LastRow(ws, refCol), refCol))) = quantiles(j) Then
                quantileRefs(j) = CLng(ws.Cells(i, 5).Value)
                Debug.Print quantileRefs(j)
                Exit For
            End If
        Next j
    Next i
    '根據輸出選項決定nthQuantile的形式
    '初始化
    NthQuantile = Array()
    '如果只要排名，直接輸出quantiles
    If outputRank Then
        NthQuantile = quantiles
    '未特別指定輸出欄，則輸出quantileRefs
    ElseIf outputCol = refCol Then
        NthQuantile = quantileRefs
    '已指定特別輸出欄位，則產生新Array
    Else
        ReDim NthQuantile(UBound(quantileRefs))
        For i = LBound(quantileRefs) To UBound(quantileRefs)
            NthQuantile(i) = ws.Cells(Application.Match(quantileRefs(i), ws.Columns(refCol), 0), outputCol).Value
        Next i
    End If

End Function

