Attribute VB_Name = "Common"
'numeric easter egg開關
Public eeTrigger As Boolean

'陣列處理用
Public Enum ArrProcMethod
    arrAppend = 1
    arrPop = 2
    arrSortDescending = 3
    arrSortAscending = 4
    arrDivide = 5
    arrSplitChar = 6
    arrInsert = 7
    arrRemove = 8
    arrReverse = 9
    arrConcat = 10
    arrSplit = 11
    arrjoin = 12
End Enum

'凍結UI選項
Public Enum UiFreezeOption
    uiDeactivate = 1
    uiActivate = 0
End Enum

Public Property Get ThisWB() As Workbook

    'thisworkbook
    Set ThisWB = ThisWorkbook
    
End Property

Public Property Get Today() As Long

    'psuedo today
    Today = DateSerial(Year(Now), Month(Now), Day(Now))

End Property

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

Public Function ArrProc(arr As Variant, Optional ipt As Variant, Optional method As ArrProcMethod = arrAppend) As Variant

    '陣列處理

    '初始化變數
    Dim workingArr As Variant
    Dim varTypeArr As Variant
    '開幹
    Select Case method
        Case arrAppend '在ubound之後再加一個元素
            ReDim Preserve arr(UBound(arr) + 1)
            arr(UBound(arr)) = ipt
            ArrProc = arr
        Case arrPop '砍掉最後一個元素，但不會回傳
            ReDim Preserve arr(UBound(arr) - 1)
            ArrProc = arr
        Case arrSortDescending, arrSortAscending 'aux array sort
            '檢查陣列有沒有需要排序
            If UBound(arr) - LBound(arr) < 2 Then
                ArrProc = arr
                Exit Function
            End If
            '決定陣列裡每個元素的型別
            varTypeArr = Array()
            For Each element In arr
                Select Case VarType(element)
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate, vbLongLong
                        vt = "dbl"
                    Case vbString
                        vt = "str"
                    Case vbVariant
                        vt = "nonsortable"
                        If IsNumeric(element) Then
                            vt = "dbl"
                        Else
                            On Error Resume Next
                            vt = CStr(element)
                            vt = "str"
                            On Error GoTo 0
                        End If
                    Case Else
                        vt = "nonsortable"
                End Select
                varTypeArr = ArrProc(varTypeArr, vt)
            Next element
            '初始化auxArr
            workingArr = Array()
            '先用ascending排數字
            For i = LBound(arr) To UBound(arr)
                If varTypeArr(i) = "dbl" Then
                    Select Case UBound(workingArr)
                        Case -1 '沒東西直接加
                            workingArr = ArrProc(workingArr, arr(i))
                        Case 0 '只有一個東西
                            If arr(i) >= workingArr(0) Then
                                workingArr = ArrProc(workingArr, arr(i))
                            Else
                                workingArr = ArrProc(workingArr, Array(0, arr(i)), arrInsert)
                            End If
                        Case Else '有兩個以上的元素
                            For j = LBound(workingArr) To UBound(workingArr)
                                compareVal = CDbl(arr(i))
                                If j = LBound(workingArr) And compareVal < workingArr(j) Then '比最小還小
                                    workingArr = ArrProc(workingArr, Array(0, arr(i)), arrInsert)
                                    Exit For
                                ElseIf j = UBound(workingArr) And compareVal >= workingArr(j) Then '比最大還大
                                    workingArr = ArrProc(workingArr, arr(i))
                                    Exit For
                                Else '其他情況
                                    If compareVal >= workingArr(j) And compareVal < workingArr(j + 1) Then
                                        workingArr = ArrProc(workingArr, Array(j + 1, arr(i)), arrInsert)
                                        Exit For
                                    End If
                                End If
                            Next j
                    End Select
                End If
            Next i
            '先把排好的數字給arrproc
            ArrProc = workingArr
            workingArr = Array()
            '再排字串
            For i = LBound(arr) To UBound(arr)
                If varTypeArr(i) = "str" Then
                    Select Case UBound(workingArr)
                        Case -1 '沒東西直接加
                            workingArr = ArrProc(workingArr, arr(i))
                        Case 0 '只有一個東西
                            If arr(i) >= workingArr(0) Then
                                workingArr = ArrProc(workingArr, arr(i))
                            Else
                                workingArr = ArrProc(workingArr, Array(0, arr(i)), arrInsert)
                            End If
                        Case Else '有兩個以上的元素
                            For j = LBound(workingArr) To UBound(workingArr)
                                compareVal = arr(i)
                                If j = LBound(workingArr) And compareVal < workingArr(j) Then '比最小還小
                                    workingArr = ArrProc(workingArr, Array(0, arr(i)), arrInsert)
                                    Exit For
                                ElseIf j = UBound(workingArr) And compareVal >= workingArr(j) Then '比最大還大
                                    workingArr = ArrProc(workingArr, arr(i))
                                    Exit For
                                Else '其他情況
                                    If compareVal >= workingArr(j) And compareVal < workingArr(j + 1) Then
                                        workingArr = ArrProc(workingArr, Array(j + 1, arr(i)), arrInsert)
                                        Exit For
                                    End If
                                End If
                            Next j
                    End Select
                End If
            Next i
            '組合arrproc跟字串排好的結果
            If UBound(ArrProc) < 0 Then '沒數字結果就帶字串結果
                ArrProc = workingArr
            ElseIf UBound(workingArr) < 0 Then '沒字串結果直接結束
                Exit Function
            Else
                ArrProc = ArrProc(ArrProc, workingArr, arrConcat)
            End If
            '如果要降冪排序就反轉
            If method = arrSortDescending Then
                ArrProc = ArrProc(ArrProc, , arrReverse)
            End If
        Case arrDivide '把arr以val元素為界切成兩個，val在右邊
            workingArr = Array(Array(), Array())
            boundary = CLng(ipt)
            '檢查能不能拆
            If boundary > UBound(arr) Then
                ArrProc = arr
                Exit Function
            End If
            '開拆
            For i = LBound(arr) To boundary - 1
                workingArr(0) = ArrProc(workingArr(0), arr(i))
            Next i
            For i = boundary To UBound(arr)
                workingArr(1) = ArrProc(workingArr(1), arr(i))
            Next i
            '回傳
            ArrProc = workingArr
        Case arrSplitChar '把str逐字拆開變arr
            workingArr = Array()
            For i = 1 To Len(arr)
                workingArr = ArrProc(workingArr, Mid(ipt, i, 1), arrAppend)
            Next i
            ArrProc = workingArr
        Case arrInsert '在val(0)前面加入元素val(1)
            '檢查val是不是array(1)跟內容
            On Error GoTo InvalidVal
            If Not IsNumeric(ipt(0)) Or (UBound(ipt) - LBound(ipt)) < 1 Then
                ArrProc = arr
                Debug.Print "喵，傳入的val不太對"
                Exit Function
            End If
            'insert位置在lb，則直接處理
            If CLng(ipt(0)) = LBound(arr) Then
                workingArr = Array()
                workingArr = ArrProc(workingArr, ipt(1))
                For i = LBound(arr) To UBound(arr)
                    workingArr = ArrProc(workingArr, arr(i))
                Next i
                ArrProc = workingArr
                Exit Function
            End If
            '其他情況先切原arr
            workingArr = ArrProc(arr, CLng(ipt(0)), arrDivide)
            arr = Array()
            '加東西進workingArr
            For i = LBound(workingArr(0)) To UBound(workingArr(0))
                arr = ArrProc(arr, workingArr(0)(i))
            Next i
            arr = ArrProc(arr, val(1))
            For i = LBound(workingArr(1)) To UBound(workingArr(1))
                arr = ArrProc(arr, workingArr(1)(i))
            Next i
            '回傳
InvalidVal:
            ArrProc = arr
        Case arrRemove '砍掉第val個元素
            'remove最後一個元素，直接跑去用pop
            If CLng(ipt) = UBound(arr) Then
                ArrProc = ArrProc(arr, , arrPop)
                Exit Function
            'remove第一個元素，直接divide取(1)
            ElseIf CLng(ipt) = LBound(arr) Then
                workingArr = ArrProc(arr, CLng(ipt) + 1, arrDivide)
                ArrProc = workingArr(1)
                Exit Function
            '其他情況就先切再pop
            Else
                workingArr = ArrProc(arr, CLng(ipt) + 1, arrDivide)
                workingArr(0) = ArrProc(workingArr(0), , arrPop)
                ArrProc = ArrProc(workingArr, , arrConcat)
                Exit Function
            End If
        Case arrReverse '反轉arr
            workingArr = Array()
            For i = UBound(arr) To LBound(arr) Step -1
                workingArr = ArrProc(workingArr, arr(i))
            Next i
            ArrProc = workingArr
        Case arrConcat '組合一個陣列，而且可用在二維陣列
            workingArr = Array()
            For i = LBound(arr) To UBound(arr)
                If VarType(arr(i)) = vbArray Then
                    For j = LBound(arr(i)) To UBound(arr(i))
                        workingArr = ArrProc(workingArr, arr(i)(j))
                    Next j
                Else
                    workingArr = ArrProc(workingArr, arr(i))
                End If
            Next i
            ArrProc = workingArr
        Case arrSplit '經典split，只用val指定delim的快速版
            If ipt = "" Then ipt = ", "
            ArrProc = Split(arr, ipt)
        Case arrjoin '經典join
            ArrProc = Join(arr, ipt)
    End Select

End Function

Public Sub UiFreeze(opt As UiFreezeOption)

    '凍結或解凍UI，進行大量計算，或潛在循環運算前後使用
    With Application
        Select Case opt
            Case uiDeactivate
                .EnableEvents = False
                .DisplayAlerts = False
                .ScreenUpdating = False
                .Calculation = xlCalculationManual
                .StatusBar = False
            Case uiActivate
                .EnableEvents = True
                .DisplayAlerts = True
                .ScreenUpdating = True
                .Calculation = xlCalculationAutomatic
                .StatusBar = True
        End Select
    End With

End Sub

Public Sub AppCalc()

    '手動呼叫計算
    With Application
        .CalculateFull
    End With
    
    Do While Application.CalculationState <> xlDone
        DoEvents
    Loop

End Sub

Public Sub AppWait(Optional waitSec As Long = 0)

    '呼叫application.wait
    Application.Wait (Now + (waitSec / 86400))

End Sub

Public Function EnumToIndex(en As Long, Optional ofst As Long = 0) As Long

    '把2指數的enum轉換成0-based index
    
    '檢查傳入的enum能不能轉
    If en <= 0 Then
        EnumToIndex = -1
        Exit Function
    End If
    
    '檢查是不是2的冪
    If en <> 2 ^ CLng(Log(en) / Log(2)) Then
        EnumToIndex = -1
        Exit Function
    End If
    
    '回傳
    EnumToIndex = CLng((Log(en) / Log(2)) + 0.49) + ofst

End Function

Public Sub OpenURL(url As String)

    '用預設瀏覽器開連結
    Shell "cmd /c start " & url, vbHide
    
End Sub

Public Function GetWs(Optional idx As wsIndex = wsNone, Optional wsName As String = "") As Worksheet

    '用enum跟property get去抓ws
    '要另外寫配對的enum wsIndex跟property get wsArr
    If wsName = "" Then
        wsName = ThisWB.Sheets(1).name
    End If

    On Error Resume Next
    If idx >= 0 Then
        Set GetWs = ThisWB.Sheets(WsArr(EnumToIndex(idx)))
    Else
        Set GetWs = ThisWB.Sheets(wsName)
    End If
    On Error GoTo 0
    
    If GetWs Is Nothing Then
        Set GetWs = ThisWB.Sheets(1)
    End If

End Function


Public Sub NumericEasterEggs(numInput As Long, Optional allowVandalism As Boolean = False)

    '一堆奇怪的彩蛋
    '慎用
    '絕對不要在認真的活頁簿裡面allowVandalism
    
    If eeTrigger Then Exit Sub
    
    eeString = ""
    eeTrigger = True

    Select Case numInput
        Case 19 'dramatic kitten
            Call OpenURL("https://www.youtube.com/shorts/JGtNdSeHkSo")
        Case 21 'happy happy happy
            reptTime = Array(3, 5, 8, 5)
            For Each t In reptTime
                For j = 1 To t
                    eeString = eeString & "happy "
                Next j
                eeString = eeString & Chr(10)
            Next t
            MsgBox eeString
        Case 123, 45680 'first day on VBA
            MsgBox "Excel Meowcro, Meows to your core since 2025"
        Case 404 'do i have to explain this
            MsgBox "Results not found. Hopefully your sense of humor does."
        Case 403 'forbidden
            MsgBox "Who said you're allowed to edit this workbook?"
            If allowVandalism Then
                MsgBox "Now I'm going to deprive you of all your progress."
                MsgBox "Bye."
                ThisWB.Close False
            Else
                MsgBox "Nah just kidding. Have a nice day."
            End If
        Case 418 'i'm a teapot
            If allowVandalism Then
                On Error Resume Next
                For Each ws In ThisWB.Worksheets
                    For Each rng In ws.UsedRange
                        If rng.Value Like "*coffee*" Then
                            rng.Clear
                        End If
                    Next rng
                Next ws
                On Error GoTo 0
            End If
        Case 1793 'louisxiv
            If allowVandalism Then
                On Error Resume Next
                For Each ws In ThisWB.Worksheets
                    With ws.PageSetup
                        .LeftHeader = ""
                        .CenterHeader = ""
                        .RightHeader = ""
                    End With
                Next ws
                On Error GoTo 0
            End If
        Case 1945 'end of wwii
            If allowVandalism Then
                On Error Resume Next
                For Each ws In ThisWB.Worksheets
                    For Each rng In ws.UsedRange
                        Select Case rng.Value
                            Case "*berlin*", "*germany*" '東西德
                                rng.Value = Replace(rng.Value, "berlin", "ber lin")
                                rng.Value = Replace(rng.Value, "germany", "ger many")
                                Exit For
                            Case "*kogun*" '皇軍解散
                                rng.Value = Replace(rng.Value, "kogun", "self-defense force")
                                Exit For
                            Case "*taiwan*" '國民政府來台
                                rng.Value = Replace(rng.Value, "taiwan", "free china")
                                Exit For
                            Case "*italy*", "*italian*" '義呆利
                                MsgBox "Go home and have your spaghetti. Don't act as if you care."
                                Exit For
                        End Select
                    Next rng
                Next ws
                On Error GoTo 0
            End If
        Case 1773 'boston tea party
            If allowVandalism Then
                On Error Resume Next
                For Each ws In ThisWB.Worksheets
                    For Each rng In ws.UsedRange
                        Select Case rng.Value
                            Case "*tea*"
                                rng.Value = Replace(rng.Value, "tea", "salt-flavored tea")
                                Exit For
                        End Select
                    Next rng
                Next ws
                On Error GoTo 0
            End If
            MsgBox "Enjoy freedom, buy lipton or fusetea and pour them into the sink."
        Case Else
            eeTrigger = False
    End Select

End Sub

