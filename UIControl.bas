Attribute VB_Name = "UIControl"
Public Sub ClearUserForm(uf As UserForm, Optional clearOptions As Boolean = True, Optional clearLabels As Boolean = False, _
                        Optional exemptName As Variant = "", Optional exemptCaption As Variant = "", _
                        Optional exemptType As Variant = "", Optional exemptionExactMatch As Boolean = False)
    
    '清空userform控件
    'clearOptions可以把listbox、combobox的選項清掉
    'clearLabels可以把label的caption清掉
    'exemptName/Caption/Type需要傳入陣列，有傳入則相對命名／字幕／類別的控件不會被清掉
    'exemptionExactMatch可以選擇要不要用萬用字元找控件
    
    Dim ctrl As Object
    
    For Each ctrl In uf.Controls
        
        isExemption = False
        
        '檢查是否略過
        If exemptName <> "" Then
            workingName = ctrl.name
            If exemptionExactMatch Then
                If Not IsError(Application.Match(workingName, exemptName, 0)) Then
                    isExemption = True
                    GoTo skipOne
                End If
            Else
                For Each nm In exemptName
                    If workingName Like "*" & nm & "*" Then
                        isExemption = True
                        GoTo skipOne
                    End If
                Next nm
            End If
        End If
        On Error Resume Next
        If exemptCaption <> "" Then
            workingCapt = ctrl.Caption
            If exemptionExactMatch Then
                If Not IsError(Application.Match(workingCapt, exemptCaption, 0)) Then
                    isExemption = True
                    GoTo skipOne
                End If
            Else
                For Each capt In exemptCaption
                    If workingCapt Like "*" & capt & "*" Then
                        isExemption = True
                        GoTo skipOne
                    End If
                Next capt
            End If
        End If
        On Error GoTo 0
        If exemptType <> "" Then
            workingType = TypeName(ctrl)
            If Not IsError(Application.Match(workingType, exemptType, 0)) Then
                isExemption = True
                GoTo skipOne
            End If
        End If
        
        '不略過則根據控件類型清空控件
        With ctrl
            Select Case TypeName(ctrl)
                Case "ListBox", "ComboBox"
                    If clearOptions Then
                        .Clear
                    End If
                    .ListIndex = -1
                Case "CheckBox", "OptionButton", "ToggleButton"
                    .Value = False
                Case "textBox"
                    .Clear
                Case "Label"
                    If clearLabels Then
                        .Caption = ""
                    End If
            End Select
        
        End With
    
skipOne:
    
    Next ctrl

End Sub

Public Sub ColWidthCalib(ws As Worksheet, Optional smart As Boolean = True, Optional bigString As Boolean = False, Optional colWidth As Variant)

    '調整欄寬
    '預設直接autofit，關閉smart後可手動傳入各欄欄寬array
    Dim i As Long
    Dim j As Long
    Dim rw As Long
    Dim cl As Long
    With ws
        If smart Then 'autofit
            '先做autofit
            cl = LastCol(ws, 1, True)
            For i = 1 To cl
                With .Columns(i)
                    .ColumnWidth = 80
                    .AutoFit
                End With
            Next i
            '有大型字串的時候再做這些
            '如果欄位裡面有換行符，或是autofit下欄寬大於50，那就直接wrap再縮（或autofit一次）
            If bigString Then
                'On Error Resume Next
                cl = LastCol(ws, 1, True)
                rw = LastRow(ws, 1, True)
                For i = 1 To cl
                    With ws.Columns(i)
                        If .ColumnWidth > 50 Then '過大調小
                            .ColumnWidth = 50
                            .AutoFit
                        End If
                    End With
                    '逐列掃描
                    For j = 1 To rw
                        With WorksheetFunction
                            If ws.Cells(j, i).Value Like "*" & Chr(10) & "*" Then '找換行符，有換行符的行距給到75
                                With ws.Columns(i)
                                    .WrapText = True
                                    .AutoFit
                                    If .ColumnWidth > 75 Then
                                        .ColumnWidth = 75
                                    End If
                                End With
                                Exit For
                            End If
                        End With
                    Next j
                Next i
                'On Error GoTo 0
            End If
        Else '固定欄寬
            For i = 0 To UBound(colWidth)
                .Columns(i + 1).ColumnWidth = colWidth(i)
            Next i
        End If
    End With

End Sub

Public Function ColWidthRetrive(ws As Worksheet, Optional scanStart As Long = 1, Optional scanEnd As Long = -1) As Variant
    
    '取得一張工作表的所有欄寬成陣列
    '可指定需要的欄位
    
    '初始化
    ColWidthRetrive = Array()
    ReDim ColWidthRetrieve(scanEnd - scanStart)
    If scanEnd < 1 Then scanEnd = LastCol(ws, 1, True)
    
    '列欄寬
    For j = 0 To scanEnd - scanStart
        ColWidthRetrieve(j) = ws.Columns(scanStart).ColumnWidth
        scanStart = scanStart + 1
    Next j
    
End Function

Public Sub DrawStandardBorders(ws As Worksheet, Optional edgeThickness As XlBorderWeight = xlThin, Optional onlyEdge As Boolean = False, Optional rng As Range, Optional smartBorders As Boolean = True)

    '畫指定範圍內黑線條的標準框線
    '可指定邊界濃度（-4138為中；4為粗）
    '可回傳特定範圍，或自動偵測邊界
    '可僅畫邊界
    
    With ws
        '決定範圍
        If rng Is Nothing Then '沒指定範圍就用smartBorders
            Set rng = .Range(.Cells(1, 1), .Cells(LastRow(ws, 1, True), LastCol(ws, 1, True)))
        End If
        '畫框線
        With rng
            '先看要不要畫內框線
            If Not onlyEdge Then
                With .Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
            '畫邊界
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = edgeThickness
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = edgeThickness
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = edgeThickness
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = edgeThickness
            End With
        End With
    End With

End Sub


