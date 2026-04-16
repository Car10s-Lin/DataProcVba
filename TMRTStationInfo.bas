Attribute VB_Name = "TMRTStationInfo"
'綠線站名選擇
Public Enum GreenLineStaSel
    g0 = 2 ^ 0
    g3 = 2 ^ 1
    g4 = 2 ^ 2
    g5 = 2 ^ 3
    g6 = 2 ^ 4
    g7 = 2 ^ 5
    g8 = 2 ^ 6
    g8a = 2 ^ 7
    g9 = 2 ^ 8
    g10 = 2 ^ 9
    g10a = 2 ^ 10
    g11 = 2 ^ 11
    g12 = 2 ^ 12
    g13 = 2 ^ 13
    g14 = 2 ^ 14
    g15 = 2 ^ 15
    g16 = 2 ^ 16
    g17 = 2 ^ 17
End Enum

'站名格式選擇
Public Enum StaNameForm
    constructCode = 0
    mandShort = 1
    mandFull = 2
    mandAlias = 3
    publicCode = 4
    engFull = 5
    engShort = 6
    mandAddress = 7
End Enum

Public Property Get TmrtGreenLineStations() As Variant

    '台中捷運綠線車站站名及資訊
    grConsCode = Array("G0", "G3", "G4", "G5", "G6", "G7", "G8", "G8a", "G9", "G10", "G10a", "G11", "G12", "G13", "G14", "G15", "G16", "G17")
    grMandShort = Array("北屯", "舊社", "松竹", "四維", "崇德", "中清", "文華", "櫻花", "市府", "水安", "森林", "南屯", "豐樂", "大慶", "九張犁", "九德", "烏日", "高鐵")
    grMandFull = Array("北屯總站", "舊社", "松竹", "四維國小", "文心崇德", "文心中清", "文華高中", "文心櫻花", "市政府", "水安宮", "文心森林公園", "南屯", "豐樂公園", "大慶", "九張犁", "九德", "烏日", "高鐵臺中站")
    grMandAlias = grMandShort
    grPubCode = Array("103a", "103", "104", "105", "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "116", "117", "118", "119")
    grEngFull = Array("Beitun Main Station", "Jiushe", "Songzhu", "Sihwei Elementary School", "Wenxin Chongde", "Wenxin Zhongqing", "Wenhua Senior High School", "Wenxin Yinghua", "Taichung City Hall", "Shui-an Temple", "Wenxin Forest Park", "Nantun", "Feng-le Park", "Daqing", "Jiuzhangli", "Jiude", "Wuri", "HSR Taichung Station")
    grEngShort = Array("Beitun", "Jiushe", "Songzhu", "Siwei", "Chongde", "Zhongqing", "Wenhua", "Yinghua", "City Hall", "Shuian", "Forest Park", "Nantun", "Fengle", "Daqing", "Jiuzhangli", "Jiude", "Wuri", "HSR")
    grMandAdd = Array("臺中市北屯區敦富東街100號", "臺中市北屯區松竹路一段1250號", "臺中市北屯區北屯路458號", "臺中市北屯區文心路四段898號", "臺中市北屯區文心路四段538號", "臺中市北區文心路三段700號", "臺中市西屯區文心路三段199號", "臺中市西屯區文心路三段107之28號", "臺中市西屯區文心路二段688號", "臺中市南屯區文心路一段519號", "臺中市南屯區文心路一段259號", "臺中市南屯區五權西路二段328號", "臺中市南屯區文心南路168號", "臺中市南區建國北路一段11號", "臺中市烏日區建國路915號", "臺中市烏日區建國路639號", "臺中市烏日區建國路295號", "臺中市烏日區高鐵東一路28號")
    TmrtGreenLineStations = Array(grConsCode, grMandShort, grMandFull, grMandAlias, grPubCode, grEngFull, grEngShort, grMandAdd)

End Property

Public Function GrSta(sta As GreenLineStaSel, Optional form As StaNameForm = 6, Optional fullNameHeadOption As Boolean = False, Optional kucing As Boolean = False) As Variant

    '以車站跟格式enum產生指定車站之指定資訊
    '可透過位元疊加一次回傳數個車站的資料陣列
       
    '初始化變數
    Dim gls As Variant
    gls = TmrtGreenLineStations
    Dim workingArr As Variant
    workingArr = Array()
    '加入對應站名
    For i = LBound(gls(0)) To UBound(gls(0))
        enm = CLng(2 ^ i)
        If enm And sta Then
            workingArr = ArrProc(workingArr, i)
        End If
    Next i
    '沒抓到車站保護
    If UBound(workingArr) < 0 Then
        GrSta = ""
        Exit Function
    End If
    '處理陣列中的元素
    For i = LBound(workingArr) To UBound(workingArr)
        workingArr(i) = gls(form)(CInt(workingArr(i)))
        '加捷運xx站
        If fullNameHeadOption Then
            Select Case form
                Case 1 To 2
                    If Right(workingArr(i), 1) = "站" Then
                        workingArr(i) = Left(workingArr(i), Len(workingArr(i)) - 1)
                    End If
                    workingArr(i) = "捷運" & workingArr(i) & "站"
                Case 5 To 6
                    If Right(workingArr(i), 7) = "Station" Then
                        workingArr(i) = WorksheetFunction.Substitute(workingArr(i), " Station", "")
                    End If
                    workingArr(i) = "Taichung Metro " & workingArr(i) & " Station"
            End Select
        End If
        '喵
        '率土之濱，莫非貓臣
        If kucing Then
            Select Case form
                Case 0 To 4
                    workingArr(i) = "貓" & workingArr(i)
                Case 5 To 6
                    workingArr(i) = "Cat " & workingArr(i)
            End Select
        End If
    Next i
    '如果陣列只有一個元素，就直接回傳字串
    If UBound(workingArr) = 0 Then
        GrSta = CStr(workingArr(0))
    Else
        GrSta = workingArr
    End If

End Function

Public Function GrStaIndex(str As Variant, Optional searchWithin As StaNameForm = -1, Optional toEnum As Boolean = True) As Long
   
    '從GrSta產出的內容反推在TmrtGreenLineStations中的索引
    
    '錯誤保護
    GrStaIndex = -1
    '初始化變數
    Dim gls As Variant
    gls = TmrtGreenLineStations
    '算出搜索範圍
    If searchWithin >= 0 Then
        lb = searchWithin
        ub = searchWithin
    Else
        lb = LBound(gls)
        ub = UBound(gls)
    End If
    '開找
    For i = lb To ub
        For j = LBound(gls(i)) To UBound(gls(i))
            If CStr(str) = CStr(gls(i)(j)) Then
                GrStaIndex = j
                '轉enum
                If toEnum Then
                    GrStaIndex = 2 ^ GrStaIndex
                End If
                Exit Function
            End If
        Next j
    Next i
    
End Function
