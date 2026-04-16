Attribute VB_Name = "DataNormalization"
'八碼格式選項
Public Enum BamaFormat
    bamaMDHM = 2 ^ 0
    bamaYMD = 2 ^ 1
    bamaYMDHM = 2 ^ 2
    bamaYMDHMS = 2 ^ 3
    bamaHM = 2 ^ 4
    bamaHMS = 2 ^ 5
    bamaMinYMD = 2 ^ 6
    bamaMinYMDHM = 2 ^ 7
    bamaMinYMDHMS = 2 ^ 8
    bamaJucheYMD = bamaMinYMD 'in case we suddenly need to worship kim family statue, want to cannon some disobeying relatives, or both (?)
    bamaJucheYMDHM = bamaMinYMDHM
    bamaJucheYMDHMS = bamaMinYMDHMS
    bamaTaishoYMD = bamaMinYMD 'in case 台灣地位未定論 become realistic someday, and the old emperor revived because of that (?)
    bamaTaishoYMDHM = bamaMinYMDHM
    bamaTaishoYMDHMS = bamaMinYMDHMS
    bamaConfuciusYMD = 2 ^ 30 'in case you don't know the kitten who wrote this also majored in chinese teaching (???)
End Enum

'儲存格格式轉換用enum
Public Enum SmartDtOption
    dtSmart = 0
    dtAllDate = 1
    dtAllTime = 2
    dtAllfull = 3
End Enum

Public Function BamaToDate(ipt As Variant, Optional fmt As BamaFormat = bamaYMD) As Date

    '八碼系時間表示形式轉DT
    
    '初始化
    BamaToDate = -1
    
    '格式錯誤保護
    If Not IsNumeric(ipt) Then
        Exit Function
    Else
        ipt = CStr(ipt)
    End If
    
    '根據格式去轉
    Select Case fmt
        Case bamaMDHM
            If Len(ipt) <> 8 Then Exit Function
            BamaToDate = DateSerial(Year(Now), CInt(Left(ipt, 2)), CInt(Mid(ipt, 3, 2))) + TimeSerial(CInt(Mid(ipt, 5, 2)), CInt(Right(ipt, 2)), 0)
            Exit Function
        Case bamaYMD
            If Len(ipt) <> 8 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 4)), CInt(Mid(ipt, 5, 2)), CInt(Right(ipt, 2)))
            Exit Function
        Case bamaYMDHM
            If Len(ipt) <> 12 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 4)), CInt(Mid(ipt, 5, 2)), CInt(Mid(ipt, 7, 2))) + TimeSerial(CInt(Mid(ipt, 9, 2)), CInt(Right(ipt, 2)), 0)
            Exit Function
        Case bamaYMDHMS
            If Len(ipt) <> 14 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 4)), CInt(Mid(ipt, 5, 2)), CInt(Mid(ipt, 7, 2))) + TimeSerial(CInt(Mid(ipt, 9, 2)), CInt(Mid(ipt, 11, 2)), CInt(Right(ipt, 2)))
            Exit Function
        Case bamaHM
            If Len(ipt) <> 4 Then Exit Function
            BamaToDate = TimeSerial(CInt(Left(ipt, 2)), CInt(Right(ipt, 2)), 0)
            Exit Function
        Case bamaHMS
            If Len(ipt) <> 6 Then Exit Function
            BamaToDate = TimeSerial(CInt(Left(ipt, 2)), CInt(Mid(ipt, 3, 2)), CInt(Right(ipt, 2)))
            Exit Function
        Case bamaMinYMD
            If Len(ipt) <> 7 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 3)) + 1911, CInt(Mid(ipt, 4, 2)), CInt(Right(ipt, 2)))
            Exit Function
        Case bamaMinYMDHM
            If Len(ipt) <> 11 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 3)) + 1911, CInt(Mid(ipt, 4, 2)), CInt(Mid(ipt, 6, 2))) + TimeSerial(CInt(Mid(ipt, 8, 2)), CInt(Right(ipt, 2)), 0)
            Exit Function
        Case bamaMinYMDHMS
            If Len(ipt) <> 13 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 3)) + 1911, CInt(Mid(ipt, 4, 2)), CInt(Mid(ipt, 6, 2))) + TimeSerial(CInt(Mid(ipt, 8, 2)), CInt(Mid(ipt, 10, 2)), CInt(Right(ipt, 2)))
            Exit Function
        Case bamaConfuciusYMD
            If Len(ipt) <> 8 Then Exit Function
            BamaToDate = DateSerial(CInt(Left(ipt, 4)) - 551, CInt(Mid(ipt, 5, 2)), CInt(Right(ipt, 2)))
            Exit Function
    End Select

End Function

Public Sub SmartNumFormat(ws As Worksheet, Optional exemptCol As Variant = False, Optional dtFormatScheme As SmartDtOption = dtSmart, Optional startRow As Long = 1, Optional dtFullFormat As String = "yyyy/m/d hh:mm:ss", Optional dateFormat As String = "yyyy/m/d", Optional timeFormat As String = "hh:mm:ss")

    '自動掃描各欄位，將可以解析為日期、時間的欄設定成需要的格式。
    '可傳入例外欄的array（絕對位置），將跳過檢查。
    '預設規則為無小數點為時間，有小數點為日期+時間，小數為時間。
    '可根據需要修改scheme，將所有時間設定成統一格式。
    '可自行設定時間格式。
    
    With ws
        '計算所用檢核列
        checkRow = WorksheetFunction.RoundUp(LastRow(ws, 1, True) / 2, 0) + startRow - 1
        If checkRow <= startRow Then checkRow = startRow + 1
        '逐欄檢核
        For i = 1 To LastCol(ws, 1, True)
            '確認是否為例外欄
            If Not VarType(exemptCol) = vbBoolean Then
                For Each col In exemptCol
                    If i = col Then
                        GoTo skipCol
                    End If
                Next col
            End If
            '19700101~20991231間的數字才需要考慮格式
            If IsNumber(.Cells(checkRow, i)) And .Cells(checkRow, i) >= 25569 And .Cells(checkRow, i) < 73051 Then
                '日期時間處理
                Select Case dtFormatScheme
                    Case dtSmart '自動偵測
                        If .Cells(checkRow, i).Value Mod 1 = 0 Then '無小數點
                            .Columns(i).NumberFormat = dateFormat
                        Else '有小數點
                            If .Cells(checkRow, i).Value >= 1 Then '含日期
                                .Columns(i).NumberFormat = dtFullFormat
                            Else '不含日期
                                .Columns(i).NumberFormat = timeFormat
                            End If
                        End If
                    Case dtAllDate '全部變日期
                        .Columns(i).NumberFormat = dateFormat
                    Case dtAllTime '全部變時間
                        .Columns(i).NumberFormat = timeFormat
                    Case dtAllfull '全部變完整格式
                        .Columns(i).NumberFormat = dtFullFormat
                End Select
            End If
skipCol:
        Next i
    End With

End Sub
