Attribute VB_Name = "DataLookup"
'appmatch函數用
Public Enum AppMatchOption
    matchGreater = -1
    matchExactly = 0
    matchMinor = 1
End Enum

Public Function AppMatch(search As Variant, within As Variant, Optional matchOption As AppMatchOption = matchExactly, Optional xmatchMode As Boolean = False) As Long

    'application.match快速呼叫
    '可用xmatchMode尋找最後一個元素
    '找不到回傳-1
    '初始化變數
    Dim tempArr As Variant
    Dim arr As Variant
    tempArr = Array()
    arr = Array()
    '初始化值
    AppMatch = -1
    '處理within
    'range轉array
    If IsObject(within) Then
        If TypeOf within Is Range Then
            tempArr = within
            If UBound(tempArr) > 1 Then '欄match
                arr = tempArr
            ElseIf UBound(tempArr) = 1 Then '列match
                c = 1
                On Error GoTo selesai
                While True
                    arr = ArrProc(arr, tempArr(1, c))
                    c = c + 1
                Wend
            End If
        Else
            Exit Function
        End If
    Else
        arr = within
    End If
selesai:
    tempArr = Array()
    '處理xmatch
    If xmatchMode Then
        tempArr = arr
        arr = Array()
        For i = UBound(tempArr) To LBound(tempArr) Step -1
            arr = ArrProc(arr, tempArr(i, 1))
        Next i
    End If
    '做appmatch
    If Not IsError(Application.Match(search, arr, matchOption)) Then
        AppMatch = Application.Match(search, arr, matchOption)
        If xmatchMode Then
            AppMatch = UBound(arr) + 1 - AppMatch + 1
        End If
    End If

End Function

Public Function SmartMid(inputStr As String, Optional start As Long = 1, Optional length As Long = 0) As String

    '可反向切字串的mid
    
    '初始化length
    If length = 0 Then length = Len(inputStr)

    '錯誤保護
    On Error GoTo skipCutting
    '如果length為正往右切]
    If length > 0 Then
        SmartMid = Mid(inputStr, start, length)
        Exit Function
    '負數則往左切
    ElseIf length < 0 Then
        SmartMid = Mid(inputStr, start + length + 1, Abs(length))
        Exit Function
    End If
    On Error GoTo 0
skipCutting:
    '錯誤或0往右切到底
    SmartMid = Mid(inputStr, start)

End Function
