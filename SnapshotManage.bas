Attribute VB_Name = "SnapshotManage"
'OMIS相關活頁簿選擇
Public Enum OmisWsOption
    omisFound = 1
    omisLostPage = 2
    omisLostSnap = 4
    omisAll = omisFound Or omisLostPage Or omisLostSnap
End Enum

'OMIS相關活頁簿特性
'排序：總欄數、主鍵欄號、標題列數、預設工作表名稱
Public Enum OmisWsPropertySelection
    omisCol = 0
    omisPk = 1
    omisHeader = 2
    omisWsName = 3
End Enum

'OMIS資料匯入選項
Public Enum OmisImportOption
    omisUpdate = 0
    omisReplace = 1
    omisAppend = 2
End Enum

'比較選項
Public Enum CompareMethod
    compareGreater = 1
    compareCurrentOrExact = 0
    compareMinor = -1
End Enum

Sub ImportCsv(url As String, targetRng As Range)

    '匯入線上csv資料表到targetRng
    
    Dim tmp As String, http As Object, stm As Object

    '下載到暫存（原封不動）
    tmp = Environ$("TEMP") & "\_tmp_gs.csv"
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then Err.Raise 5, , "HTTP " & http.Status

    Set stm = CreateObject("ADODB.Stream")
    Dim bom(2) As Byte
    bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF
    With stm
        .Type = 1: .Open
        .Write bom
        .Write http.responseBody           '不改任何內容（binary 寫入）
        .SaveToFile tmp, 2                 'adSaveCreateOverWrite
        .Close
    End With

    '用 OpenText 的解析器（能吃引號內多重換行）
    Application.ScreenUpdating = False
    Workbooks.Open Filename:=tmp, ReadOnly:=True

    '複製到目標，關掉暫存活頁簿
    With ActiveWorkbook
        .Sheets(1).UsedRange.Copy targetRng
        .Close SaveChanges:=False
    End With
    
    On Error Resume Next: Kill tmp: On Error GoTo 0
    Application.ScreenUpdating = True
    
End Sub

Public Function GSheetLinkParser(wbId As String, wsId As String, Optional smartWbId As Boolean = True, Optional smartWsId As Boolean = True)

    '做出ghseet下載csv的連結
    '預設smartWbId=true，從完整連結直接抓活頁簿ID
    '預設smartWsId=true，用工作表名稱去抓gid
    '取得wsId
    If smartWsId Then
        wsId = WorksheetFunction.EncodeURL(wsId)
    End If
    '剪出wbId
    If smartWbId Then
        wbId = Mid(wbId, WorksheetFunction.search("/d/", wbId) + 3, WorksheetFunction.search("/edit", wbId) - WorksheetFunction.search("/d/", wbId) - 3)
    End If
    '取得連結
    Dim gUrl As String
    If smartWsId Then
        gUrl = "https://docs.google.com/spreadsheets/d/" & wbId & "/gviz/tq?tqx=out:csv&sheet=" & wsId
    Else
        gUrl = "https://docs.google.com/spreadsheets/d/" & wbId & "/export?format=csv&gid=" & wsId
    End If
    '產出結果
    GSheetLinkParser = gUrl

End Function

Public Function ListFiles(Optional multiSel As Boolean = True, Optional fileType As Variant = "", Optional fdgTitle As String = "選擇檔案喵") As Collection
    
    '透過filedialog收集一個或數個檔案路徑到collection，以利後續處理
    '可以指定AllowMultiSelection、檔案類型（類型跟顯示名稱共用字串）、filedialog的標題
    
    '初始化
    Dim workingList As New Collection
    If fileType = "" Then
        fileType = Array("*.xls*")
    End If
    
    '開檔案之後把路徑列成一個collection供處理
    '可以指定multiselect、filter（傳入array)、Title
    Dim fdg As FileDialog
    Dim file As Variant
    
    '放filedialog、整理檔案清單
    Set fdg = Application.FileDialog(msoFileDialogFilePicker)
    
    With fdg
        .AllowMultiSelect = multiSel
        .title = fdgTitle
        With .Filters
            .Clear
            For Each tp In fileType
                .Add tp, tp
            Next tp
        End With
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = -1 Then
            For Each file In .SelectedItems
                workingList.Add file
            Next file
        Else
            MsgBox "喵，沒選到檔案"
            Exit Function
        End If
    End With
    
    '回傳
    Set ListFiles = workingList
        
End Function

Public Sub OmisDataImporter(Optional wsOpt As OmisWsOption = omisAll, Optional importOpt As OmisImportOption = omisUpdate, _
                            Optional pdp As Boolean = True, Optional autoBookType As Boolean = True, Optional foundWs As Worksheet, _
                            Optional lostPageWs As Worksheet, Optional lostSnapWs As Worksheet)

    '匯入OMIS資料用
    '僅為ListFiles的使用範例，其中包含特定資料格式的匯入方法，無法一體適用所有資料表
    '隱去關鍵內部作業邏輯
    
    '初始化變數
    Dim workingWsArr As Variant
    workingWsArr = Array()
    Dim omisWsArr As Variant
    omisWsArr = Array(omisLostPage, omisFound, omisLostSnap)
    Dim fullFileList As New Collection
    Dim workingFileList As New Collection
    Dim workingWs As Worksheet
    Dim sauceWb As Workbook
    Dim workingSauce As Worksheet
    '關UI
    Call UiFreeze(uiDeactivate)
    '決定scope of work
    needListing = False
    For Each o In omisWsArr
        If o And wsOpt Then
            workingWsArr = ArrProc(workingWsArr, o)
            Select Case o
                Case omisFound, omisLostSnap
                    needListing = True
            End Select
        End If
    Next o
    '抓檔案
    If autoBookType And needListing Then
        fullFileList = ListFiles(True, "", "選擇所有所需檔案")
    End If
    '決定目標ws
    For i = LBound(OmisWsProperty) To UBound(OmisWsProperty)
        Select Case i
            Case 0
                Set workingWs = foundWs
            Case 1
                Set workingWs = lostPageWs
            Case 2
                Set workingWs = lostSnapWs
        End Select
        If workingWs Is Nothing Then
            On Error Resume Next
            Select Case i
                Case 0
                    Set foundWs = ThisWB.Sheets(OmisWsProperty(i)(omisWsName))
                Case 1
                    Set lostPageWs = ThisWB.Sheets(OmisWsProperty(i)(omisWsName))
                Case 2
                    Set lostSnapWs = ThisWB.Sheets(OmisWsProperty(i)(omisWsName))
            End Select
            On Error GoTo 0
        End If
    Next i
    '逐一處理需求
    For Each o In workingWsArr
        Select Case o
            Case omisLostPage '網頁複製貼上的那份
                '定義目標工作表
                Set workingWs = Nothing
                Set workingWs = lostPageWs
                '無目標保護
                If workingWs Is Nothing Then GoTo nextO
                'replace先清資料
                If importOpt = omisReplace Then
                    Call OmisDataPurger(omisLostPage, , workingWs, , True)
                End If
                '清掉隱藏的物件
                On Error Resume Next
                With workingWs
                    For Each ctrl In .OLEObjects
                        ctrl.Delete
                    Next ctrl
                    For Each shp In .Shapes
                        shp.Delete
                    Next shp
                End With
                On Error GoTo 0
                '整理空白欄、重複標題列
                lr = LastRow(workingWs, , True)
                For i = lr To 2 Step -1
                    With workingWs
                        If .Cells(i, 1).Value = "" Then
                            .Range(.Cells(i, 1), .Cells(i, 1)).Delete xlShiftToLeft
                        ElseIf .Cells(i, 1).Value = "item_record_id" Then
                            .Range(.Cells(i, 1), .Cells(i, 10)).Delete xlShiftUp
                        End If
                    End With
                Next i
                '維持單一值
                If importOpt = omisUpdate Then
                    Call SoloVal(workingWs, NameToColIndex("item_record_id", workingWs))
                End If
            Case Else '匯入資料表的那些
                '定義目標工作表
                Set workingWs = Nothing
                Select Case o
                    Case omisFound
                        Set workingWs = foundWs
                        wbKw = "拾得資料"
                        wsName = "工作表1"
                    Case omisLostSnap
                        Set workingWs = lostSnapWs
                        wbKw = "協尋資料"""
                        wsName = "Sheet1"
                End Select
                '無目標保護
                If workingWs Is Nothing Then GoTo nextO
                '獲取來源活頁簿
                Set workingFileList = New Collection
                If autoBookType Then
                    For Each wbLoc In fullFileList
                        If wbLoc Like "*" & wbKw & "*" Then
                            workingFileList.Add wbLoc
                        End If
                    Next wbLoc
                Else
                    workingFileList = ListFiles(True, "", "選擇" & OmisWsProperty(EnumToIndex(o))(omisWsName) & "喵")
                End If
                '無來源保護
                If workingFileList.Count = 0 Then GoTo nextO
                '逐一處理來源
                For Each wb In workingFileList
                    '取得工作表
                    Set sauceWb = Workbooks.Open(wb, ReadOnly:=True)
                    Set workingSauce = sauceWb.Sheets(wsName)
                    '統計列數
                    rws = LastRow(workingSauce, OmisWsProperty(EnumToIndex(o))(omisPk)) - OmisWsProperty(EnumToIndex(o))(omisHeader)
                    '在目標插入相對列數
                    With workingWs
                        For i = 1 To rws
                            .Rows(2).Insert
                        Next i
                    End With
                    '複製
                    With workingSauce
                        .Range(.Cells(OmisWsProperty(EnumToIndex(o))(omisHeader) + 1, 1), .Cells(LastRow(workingSauce, OmisWsProperty(EnumToIndex(o))(omisPk)), OmisWsProperty(EnumToIndex(o))(omisCol))).Copy
                    End With
                    'pastespecial
                    With workingWs
                        .Range(.Cells(2, 1), .Cells(2, 1)).PasteSpecial xlPasteValuesAndNumberFormats
                    End With
                    '關wb
                    sauceWb.Close False
                Next wb
                '維持單一值
                If importOpt = omisUpdate Then
                    Call SoloVal(workingWs, OmisWsProperty(EnumToIndex(o))(omisPk))
                End If
        End Select
        '個人資料保護措施
        Call PersonalDataProtection(workingWs, , , OmisWsProperty(EnumToIndex(o))(omisHeader))
nextO:
    Next o
    '開UI喵
    Call UiFreeze(uiActive)

End Sub

Public Sub OmisDataPurger(opt As OmisWsOption, Optional foundWs As Worksheet, Optional lostPageWs As Worksheet, _
                            Optional lostSnapWs As Worksheet, Optional clearHeader As Boolean = False)

    '清空OMIS資料
    '搭配OmisDataImporter設計的資料清除sub範例，已隱去內部作業關鍵邏輯
    
    '定義需要清除的欄位、索引欄
    Dim allOpt As Variant
    Dim workingOpt As Variant
    allOpt = Array(omisFound, omisLostPage, omisLostSnap)
    workingOpt = Array()
    For Each o In allOpt
        If o And opt Then
            workingOpt = ArrProc(workingOpt, o)
        End If
    Next o
    '空陣列保護
    If UBound(workingOpt) < 0 Then Exit Sub
    '清理門戶（？）
    Dim workingWs As Worksheet
    For Each o In workingOpt
        '復歸workingWs
        Set workingWs = Nothing
        '取得o的位置
        itmIdx = EnumToIndex(o)
        '取得索引、欄數、預設工作表
        cols = OmisWsProperty(itmIdx)(omisCol)
        idx = OmisWsProperty(itmIdx)(omisPk)
        On Error Resume Next
        Set workingWs = ThisWB.Sheets(OmisWsProperty(itmIdx)(omisWsName))
        On Error GoTo 0
        '取得特殊工作表
        If workingWs Is Nothing Then
            Select Case o
                Case omisFound
                    If Not foundWs Is Nothing Then
                        Set workingWs = foundWs
                    End If
                Case omisLostPage
                    If Not lostPageWs Is Nothing Then
                        Set workingWs = lostPageWs
                    End If
                Case omisLostSnap
                    If Not lostSnapWs Is Nothing Then
                        Set workingWs = lostSnapWs
                    End If
            End Select
        End If
        '無工作表保護
        If workingWs Is Nothing Then GoTo NextWs
        '決定要不要清標題
        hdr = 0
        If clearHeader Then hdr = OmisWsProperty(itmIdx)(omisHeader)
        '清
        Call ClearWs(workingWs, hdr, 0, idx)
NextWs:
    Next o

End Sub

Public Sub ClearWs(ws As Worksheet, Optional hearerRows As Long = 0, Optional headerCols As Long = 0, Optional indexRow As Long = 1, Optional indexCol As Long = 1, Optional chaoMode As Boolean = False)

    '清除整張工作表的資料
    '可指定特定標題攔或列為標題欄列，不予清除
    With ws
        .Range(.Cells(1 + headerRows, 1 + headerCols), .Cells(LastRow(ws, indexCol, chaoMode), LastCol(ws, indexRow, chaoMode))).Clear
    End With

End Sub

Public Sub PersonalDataProtection(ws As Worksheet, Optional keywords As Variant, Optional customCols As Variant, Optional colNameRow As Long = 1, Optional keepDataTypeAndLen As Boolean = True)

    '掃描資料中包含個資的欄位，並加以刪除或隱蔽（僅保留型別及長度）
    '可指定欄位或自動掃描標題列
    '初始化變數
    Dim workingCols As Variant
    workingCols = Array()
    Dim kwArr As Variant
    kwArr = Array("int", "float", "str", "erased")
    '決定隱去欄位的方式
    If customCols Is Nothing Then '用關鍵字去比對
        If keywords Is Nothing Then '未指定關鍵字，使用預設關鍵字
            keywords = Array("姓名", "身分證字號", "身分證號", "電話", "手機")
        End If
        With ws
            For i = 1 To LastCol(ws, colNameRow)
                chkVal = .Cells(colNameRow, i).Value
                For Each kw In keywords
                    If chkVal Like "*" & kw & "*" Then
                        workingCols = ArrProc(workingCols, i)
                        Exit For
                    End If
                Next kw
            Next i
        End With
    Else '指定欄位
        '偵測是不是用欄序傳入
        isColNum = True
        For i = LBound(customCols) To UBound(customCols)
            If Not IsNumeric(customCols) Then
                isColNum = False
                Exit For
            End If
        Next i
        '如果不是欄序就代換成欄序
        If Not isColNum Then
            For i = LBound(customCols) To UBound(customCols)
                workingCols = ArrProc(workingCols, NameToColIndex(customCols(i), ws, , , colNameRow))
            Next i
        End If
    End If
    '逐個代換workingCols裡面的所有值
    With ws
        For i = colNameRow + 1 To LastRow(ws, , True)
            For Each col In workingCols
                If .Cells(i, col).Value <> "" Then
                    orgStr = .Cells(i, col).Value
                    workingStr = ""
                    '確認是否已代換過
                    isErased = False
                    For Each kw In kwArr
                        If orgStr Like "*" & kw & "*" Then
                            isErased = True
                            Exit For
                        End If
                    Next kw
                    If Not isErased Then
                        If keepDataTypeAndLen Then '保留型別及長度
                            '型別判定
                            If IsNumeric(orgStr) Then '數字
                                If CDbl(orgStr) Mod 1 = 0 Then '整數
                                    workingStr = workingStr & "int"
                                Else '分小數
                                    workingStr = workingStr & "float"
                                End If
                            Else '字串
                                workingStr = workingStr & "str"
                            End If
                            '長度判定
                            workingStr = workingStr & "(" & CStr(Len(orgStr)) & ")"
                            .Cells(i, col).Value = workingStr
                        Else  '不保留型別及長度
                            .Cells(i, col).Value = "erased"
                        End If
                    End If
                End If
            Next col
        Next i
    End With
                        
End Sub

Public Sub SoloVal(ws As Worksheet, pkCol As Long, Optional startRow As Long = 1, Optional compareCol As Long = 0, Optional hapusApa As CompareMethod = compareMinor)

    '根據主鍵（或其他特定欄位）保留單一值
    '可選擇資料起始列（不含標題列）
    
    '初始化變數
    delRowSentinel = "kucing kecil bahagia mau hapus ini, hapus manual kapan berada"
    If compareCol = 0 Then compareCol = pkCol
    
    With ws
    
        For i = LastRow(ws, pkCol) To startRow Step -1 '遍歷所有列，找需要刪除的列
            curPk = .Cells(i, pkCol).Value '取得目前pk
            For j = i - 1 To startRow Step -1
                If curPk = .Cells(j, pkCol).Value Then '找到同樣的pk
                    If pkCol = compareCol Then '如果沒要比較，刪除找到的那個
                        .Cells(j, pkCol).Value = delRowSentinel
                    Else
                        curVal = .Cells(i, compareCol).Value
                        foundVal = .Cells(j, compareCol).Value
                        Select Case hapusApa
                            Case compareGreater '刪掉數值最大的
                                If foundVal > curVal Then
                                    .Cells(j, pkCol).Value = delRowSentinel
                                Else
                                    .Cells(i, pkCol).Value = delRowSentinel
                                End If
                            Case compareMinor '刪掉數值最小的
                                If foundVal < curVal Then
                                    .Cells(j, pkCol).Value = delRowSentinel
                                Else
                                    .Cells(i, pkCol).Value = delRowSentinel
                                End If
                            Case compareCurrentOrExact '沒說要刪什麼，刪除找到的那個
                                .Cells(j, pkCol).Value = delRowSentinel
                        End Select
                    End If
                    Exit For
                End If
            Next j
        Next i
        
        For i = LastRow(ws, pkCol) To startRow Step -1 '開刪
            If .Cells(i, pkCol).Value = delRowSentinel Then
                .Rows(i).Delete
            End If
        Next i
        
    End With
    
End Sub
