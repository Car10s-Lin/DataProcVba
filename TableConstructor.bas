Attribute VB_Name = "TableConstructor"
'產sn選項
Public Enum GetSnOption
    snBesar = 1
    snNew = 2
    snKecil = 3
    snVacancy = 4
    snRand = 5
End Enum

Public Sub FillTitle(ws As Worksheet, title As Variant, Optional offseting As Long = 0, Optional fillRow As Boolean = False, Optional fillBegin As Long = 1, Optional fillFin As Long = 1048576)

    '根據title array逐欄填入標題
    '可改為逐列（fillRow = true)
    '也可指定不填在第一列／欄（offseting）
    '或是從特定欄／列開始填（fillBegin、fillFin）
    
    With ws
        For i = fillBegin To WorksheetFunction.Min(fillFin, UBound(title) + 1)
            If fillRow Then '逐列填
                .Cells(i, 1 + offseting).Value = title(i - 1)
            Else '逐欄填
                .Cells(1 + offseting, i).Value = title(i - 1)
            End If
        Next i
    End With

End Sub

Public Sub GenSequence(ws As Worksheet, Optional startRow As Long = 2, Optional col As Long = 1, Optional startNum As Long = 1, Optional gap As Long = 1)

    '產生等差級數的序列號
    
    With ws
        For i = startRow To LastRow(ws, 1, True)
            .Cells(i, col) = startNum
            startNum = startNum + gap
        Next i
    End With

End Sub

Public Function GetSn(ws As Worksheet, prefix As String, Optional col As Long = 1, Optional opt As GetSnOption = snNew, Optional snLen As Long = 7, Optional smartSnLen As Boolean = False) As String

    '根據prefix掃描特定工作表內的sn，並根據規則找出特定sn
    '規則包含最大、最小、最大+1、最快找到的可用、隨機
    
    '初始化變數
    Dim existSnArr As Variant
    existSnArr = Array()
    If smartSnLen Then snLen = 0
    workingSn = 1
    '找出已經用過的sn
    With ws
        lr = LastRow(ws, col)
        For i = 1 To lr
            workingStr = .Cells(i, col).Value
            If workingStr Like prefix & "*" Then
                isSn = False
                isSn = IsNumeric(Replace(workingStr, prefix, ""))
                If isSn Then
                    existSnArr = ArrProc(existSnArr, CLng(Replace(workingStr, prefix, "")))
                    '如果要自動判定後續需要的長度，在這裡檢驗sn的len
                    If smartSnLen Then
                        workingLen = Len(workingStr) - Len(prefix)
                        If workingLen > snLen Then
                            snLen = workingLen
                        End If
                    End If
                End If
            End If
        Next i
    End With
    '無長度保護
    If snLen < 1 Then snLen = 7
    '無已知序號保護
    If UBound(existSnArr) < 0 Then GoTo skipCalc
    '根據規則找出需要的sn
    Select Case opt
        Case snBesar, snNew '找最大或最大+1
            workingSn = WorksheetFunction.Max(existSnArr)
            If opt = snNew Then workingSn = workingSn + 1
        Case snKecil '找最小
            workingSn = WorksheetFunction.Min(existSnArr)
        Case snVacancy '找隨便一個可用
            For i = LBound(existSnArr) To UBound(existSnArr)
                If i < UBound(existSnArr) Then
                    num1 = existSnArr(i + 1)
                    num2 = existSnArr(i)
                    If Abs(num1 - num2) > 1 Then
                        matchResult = 1048576
                        If num1 > num2 Then
                            For j = num2 + 1 To num1 - 1
                                matchResult = AppMatch(j, existSnArr)
                                If matchResult < 0 Then
                                    workingSn = j
                                    Exit For
                                End If
                            Next j
                        Else
                            For j = num1 + 1 To num2 - 1
                                matchResult = AppMatch(j, existSnArr)
                                If matchResult < 0 Then
                                    workingSn = j
                                    Exit For
                                End If
                            Next j
                        End If
                        If matchResult < 0 Then
                            Exit For
                        End If
                    End If
                Else
                    workingSn = existSnArr(i)
                    isFound = False
                    Do While Not isFound
                        workingSn = workingSn + 1
                        matchResult = AppMatch(workingSn, existSnArr)
                        If matchResult < 0 Then
                            isFound = True
                        End If
                    Loop
                End If
            Next i
        Case snRand '找隨機
            isFound = False
            Do While Not isFound
                workingSn = WorksheetFunction.RandBetween(1, (10 ^ snLen) - 1)
                matchResult = AppMatch(workingSn, existSnArr)
                If matchResult < 0 Then
                    isFound = True
                End If
            Loop
    End Select
skipCalc:
    '產出序號
    fmt = ""
    For i = 1 To snLen
        fmt = fmt & "0"
    Next i
    GetSn = prefix & Format(workingSn, fmt)

End Function
