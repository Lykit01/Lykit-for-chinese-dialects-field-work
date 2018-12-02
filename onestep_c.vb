'20180824
'更新日志：零音1.0，整合版，无中古地位
Sub onestep_c()
    Dim inputStr As String '在这里输入一些必要的信息, 在不含声调不重复的单音节中，当某个声韵调小于等于excepVal时被认为是特例
    inputStr = InputBox(prompt:="格式：声，韵，调，如1,1,10。不输入的话默认为0,0,0", Title:="请输入区分声韵调特殊项的阈值", Default:="0,0,0")
    If inputStr Like "*,*,*" Then
            inputArr = Split(inputStr, ",")
            If UBound(inputArr) > 2 Then
                MsgBox "输入格式有误,默认为0,0,0"
                Exit Sub
                Else
                initExcepVal = Val(inputArr(0))
                fnlExcepVal = Val(inputArr(1))
                toneExcepVal = Val(inputArr(2))
            End If
        Else
            MsgBox "输入格式有误,默认为0,0,0"
            Exit Sub
    End If
    
    Dim beginMoment '记录程序耗时
    beginMoment = Timer
    
    Dim i As Integer, k As Integer, j As Integer, n As Integer
    Dim tgSht As Integer '记音的音节放置的位置
    tgSht = 1 'excel里放目标sheet的位置
    Dim obArr
    obArr = Sheets(tgSht).Range("a1").CurrentRegion.Value
    Dim colTitle As String '汉字的列名以及判断是汉字列的标准
    Dim hzCol As Integer, orgSylCol As Integer, wenbaiCol As Integer, shiyiCol As Integer, beizhuCol As Integer '文白、释义和备注的位置
    For i = 1 To UBound(obArr, 2)
        colTitle = obArr(1, i)
        If in_arr(Split("汉字，词项，字，词，词目", "，"), colTitle) > 0 Then hzCol = i
        If colTitle Like "*记音*" Then orgSylCol = i
        If colTitle Like "*文白*" Then wenbaiCol = i
        If colTitle Like "*义*" Then shiyiCol = i
        If colTitle Like "*备注*" Then beizhuCol = i
    Next
    
    Dim monSylCount As Integer, polySylCount As Integer
    Dim obSectCol As Integer, sylCol As Integer, obInitRowsCount As Integer '切分结果放置的位置\初始行数
    sylCol = UBound(obArr, 2) + 1
    obSectCol = UBound(obArr, 2) + 2
    obInitRowsCount = UBound(obArr, 1)
    ReDim Preserve obArr(1 To UBound(obArr, 1), 1 To UBound(obArr, 2) + 9)
    obColTitleArr = Array("音节修正", "声韵组合", "声母", "韵母", "声调", "韵尾", "标记", "手工检查", "polyphone")
    For i = 0 To UBound(obColTitleArr)
        obArr(1, sylCol + i) = obColTitleArr(i)
    Next
    '去重
    Dim dedup() As String, dupRec As Integer
    ReDim dedup(2 To UBound(obArr, 1))
    For i = 2 To UBound(obArr, 1)
        dedup(i) = obArr(i, hzCol) + obArr(i, orgSylCol)
    Next
    dupSortedPos = range_arr(2, UBound(obArr, 1), 2)
    Call quick_sort(dedup, dupSortedPos, 2, UBound(dedup))
    i = 3
    Do While i <= UBound(obArr, 1)
        If dedup(i) = dedup(i - 1) Then
            dupRec = dupRec + 1
            Else
                dupRec = 0
        End If
        If dupRec = 1 Then
            obArr(dupSortedPos(i - 1), obSectCol + 6) = "重复留存"
            obArr(dupSortedPos(i), obSectCol + 6) = "重复"
            ElseIf dupRec > 1 Then
                obArr(dupSortedPos(i), obSectCol + 6) = "重复"
        End If
        i = i + 1
    Loop
    Dim pp() As String, ppRec As Integer, ppNum As Integer, ppTag As Integer, diffPpCount As Integer 'polyphone,一字多音标记
    ReDim pp(2 To UBound(obArr, 1))
    '切分多音节与单、多音节上标记
    Dim syl As String, obRowsTmp As Integer
    Dim polySylArr() As String
    obArr = array_transpose(obArr) '第一维要扩充，先转置
    For i = 2 To UBound(obArr, 2)
        syl = std_syl(obArr(orgSylCol, i)) '规范记音格式
        If obArr(obSectCol + 6, i) <> "重复" Then obArr(sylCol, i) = syl '修正过后的音节放在此处,重复的音节不要，同时第一个重复的音节是要的
        If obArr(orgSylCol, i) <> syl Then obArr(obSectCol + 6, i) = obArr(obSectCol + 6, i) + "不规范"
        syl = obArr(sylCol, i) '重复的音节不要
        If syl <> "" Then
            pp(i) = obArr(hzCol, i) 'polyphone,一字多音标记
            If InStr(syl, " ") <> 0 Then '判断单/多音节
                obArr(obSectCol + 5, i) = "多音节"
                polySylCount = polySylCount + 1 '记录多音节的个数
                ReDim Preserve polySylArr(1 To polySylCount) '记录多音节的坐标
                polySylArr(polySylCount) = i
                polySylSectArr = Split(syl, " ") '把多音节拆分成单音节
                obRowsTmp = UBound(obArr, 2)
                ReDim Preserve obArr(1 To UBound(obArr, 1), 1 To UBound(obArr, 2) + UBound(polySylSectArr) + 1)
                For j = 0 To UBound(polySylSectArr)
                    obArr(hzCol, obRowsTmp + j + 1) = obArr(hzCol, i) + CStr(j + 1)
                    obArr(sylCol, obRowsTmp + j + 1) = polySylSectArr(j)
                Next
                Else
                    obArr(obSectCol + 5, i) = "单音节"
                    monSylCount = monSylCount + 1 '记录单音节的个数
            End If
        End If
    Next
    obArr = array_transpose(obArr) '转置回来
    ppSortedPos = range_arr(2, UBound(obArr, 1), 2) 'polyphone,一字多音标记
    Call quick_sort(pp, ppSortedPos, 2, UBound(pp))
    i = 3
    ppRec = 1
    Do While i <= UBound(pp)
        If pp(i) = pp(i - 1) Then
            ppRec = ppRec + 1
            ppNum = 0
            Else
                If ppRec >= 2 Then ppNum = ppRec
                ppRec = 1
        End If
        ppTag = ppNum
        If ppTag <> 0 Then diffPpCount = diffPpCount + 1
        Do While ppNum >= 1
            obArr(ppSortedPos(i - ppNum), obSectCol + 7) = "pp" + CStr(ppTag)
            ppNum = ppNum - 1
        Loop
        If pp(i) = "" Then Exit Do
        i = i + 1
    Loop 'polyphone,一字多音标记
    '这些是装坐标和内容的数组
    Dim sylRowArr() As Integer '不重复的整个音节的行位置
    ReDim sylRowArr(1 To 1)
    Dim sylCountArr() As Integer
    ReDim sylCountArr(1 To 1)
    
    Dim tonelessSylArr() As String '声、韵组合的行位置
    ReDim tonelessSylArr(1 To 1)
    Dim tonelessSylCountArr() As Integer
    ReDim tonelessSylCountArr(1 To 1)
    
    Dim initArr() As String
    ReDim initArr(1 To 1)
    Dim initCountArr() As Integer
    ReDim initCountArr(1 To 1)
    
    Dim fnlArr() As String
    ReDim fnlArr(1 To 1)
    Dim fnlCountArr() As Integer
    ReDim fnlCountArr(1 To 1)
    
    Dim toneArr() As String
    ReDim toneArr(1 To 1)
    Dim toneCountArr() As Integer
    ReDim toneCountArr(1 To 1)

    Dim endArr() As String
    ReDim endArr(1 To 1)
    Dim endCountArr() As Integer
    ReDim endCountArr(1 To 1)
    '预先读取第一条数据的值，防止程序在比较时报错
    sylRowArr(1) = 2
    Do While obArr(sylRowArr(1), sylCol) = ""
        sylRowArr(1) = sylRowArr(1) + 1
    Loop
    sylSectArr = syl_sect(obArr(sylRowArr(1), sylCol))
    tonelessSylArr(1) = sylSectArr(0)
    initArr(1) = sylSectArr(1)
    fnlArr(1) = sylSectArr(2)
    toneArr(1) = sylSectArr(3)
    endArr(1) = sylSectArr(4)
    '切分声韵调与计数
    Dim notSame As Boolean
    For i = sylRowArr(1) To UBound(obArr, 1) '从sylRowArr(1)开始，跳过没有记音的行数
        syl = obArr(i, sylCol)
        If syl <> "" And InStr(syl, " ") = 0 Then
            sylSectArr = syl_sect(syl)
            For k = 0 To UBound(sylSectArr)
                obArr(i, obSectCol + k) = sylSectArr(k)
            Next
            '含声调含特殊项的所有不相同的单音节,sylRowArr记录的其实是这些音节在obArr中的地址
            notSame = True
            For j = 1 To UBound(sylRowArr)
                If obArr(sylRowArr(j), obSectCol) + obArr(sylRowArr(j), obSectCol + 3) = sylSectArr(0) + sylSectArr(3) Then   '因为可能有黏着音节，所以采取纯声韵调的相加
                    notSame = False
                    sylCountArr(j) = sylCountArr(j) + 1
                    Exit For '找到相同的就可以跳出for循环了
                End If
            Next
            If notSame Then
                ReDim Preserve sylRowArr(1 To UBound(sylRowArr) + 1)
                ReDim Preserve sylCountArr(1 To UBound(sylCountArr) + 1)
                sylRowArr(UBound(sylRowArr)) = i
                sylCountArr(UBound(sylCountArr)) = 1
            End If
            
            Call catg_count(tonelessSylArr, tonelessSylCountArr, sylSectArr(0))
            Call catg_count(initArr, initCountArr, sylSectArr(1))
            Call catg_count(fnlArr, fnlCountArr, sylSectArr(2))
            Call catg_count(toneArr, toneCountArr, sylSectArr(3))
            Call catg_count(endArr, endCountArr, sylSectArr(4))
        End If
    Next
    
    Call quick_sort(initCountArr, initArr, 1, UBound(initCountArr))
    Call quick_sort(fnlCountArr, fnlArr, 1, UBound(fnlCountArr))
    Call quick_sort(toneCountArr, toneArr, 1, UBound(toneCountArr))
    Call quick_sort(endCountArr, endArr, 1, UBound(endCountArr))
    Call quick_sort(tonelessSylCountArr, tonelessSylArr, 1, UBound(tonelessSylCountArr))
    '出现次数多的声韵调，可以归入音系中
    Dim toneType()
    Dim initType()
    Dim fnlType()
    Call find_type(toneType, toneCountArr, toneArr, toneExcepVal)
    Call find_type(initType, initCountArr, initArr, initExcepVal)
    Call find_type(fnlType, fnlCountArr, fnlArr, fnlExcepVal)
    
    '同音字表
    '下面是按照语音学知识将声母、韵母和声调排序
    Dim ssToneCount As Integer, csToneCount As Integer, ssFnlCount As Integer, csFnlCount As Integer
    Call init_phon_sort(initType)
    initSysTbl = init_tbl(initType)
    Call fnl_phon_sort(fnlType)
    fnlTmp = fnl_tbl(fnlType)
    fnlSysTbl = fnlTmp(0)
    csFnlCount = fnlTmp(1)
    ssFnlCount = UBound(fnlType) - csFnlCount
    Call tone_phon_sort(toneType)
    
    If InStr(toneType(1), "入") > 0 Then '在生成韵母时要区分入声韵和舒声韵的声调
        csToneCount = UBound(toneType)
        ssToneCount = 0
        ElseIf InStr(toneType(UBound(toneType)), "入") = 0 Then
            ssToneCount = UBound(toneType)
            csToneCount = 0
        Else
            Dim ssToneType()
            Dim csToneType()
            For i = 1 To UBound(toneType)
                If InStr(toneType(i), "入") > 0 Then
                    csToneCount = csToneCount + 1
                    ReDim Preserve csToneType(1 To csToneCount)
                    csToneType(csToneCount) = toneType(i)
                    Else
                        ssToneCount = ssToneCount + 1
                        ReDim Preserve ssToneType(1 To ssToneCount)
                        ssToneType(ssToneCount) = toneType(i)
                End If
            Next
    End If
    If csToneCount = UBound(toneType) Then csToneType = toneType
    If ssToneCount = UBound(toneType) Then ssToneType = toneType
    
    Dim hpShtArr() As String
    ReDim hpShtArr(1 To csFnlCount * csToneCount + ssFnlCount * ssToneCount + 1, 1 To UBound(initType) + 2)
    Dim eaHpCount() As Integer
    ReDim eaHpCount(1 To csFnlCount * csToneCount + ssFnlCount * ssToneCount + 1, 1 To UBound(initType) + 2)
    Dim hpTbl() As String
    ReDim hpTbl(1 To UBound(initType), 1 To UBound(fnlType), 1 To UBound(toneType))
    Dim hpTblCnt() As Integer
    ReDim hpTblCnt(1 To UBound(initType), 1 To UBound(fnlType), 1 To UBound(toneType))
   
    For i = 1 To UBound(initType) '同音字表横和纵的栏目
        hpShtArr(1, i + 2) = initType(i)
    Next
    Dim hpRowRec As Integer, notCs As Boolean
    hpRowRec = 2
    Dim hpFnlRowArr()
    ReDim hpFnlRowArr(1 To UBound(fnlType))
    For i = 1 To UBound(fnlType)
        hpFnlRowArr(i) = hpRowRec
        hpShtArr(hpRowRec, 1) = fnlType(i)
        notCs = True
        If InStr("?ptk", right(fnlType(i), 1)) > 0 Then notCs = False
        If notCs Then
            For j = 1 To UBound(ssToneType)
                hpShtArr(j - 1 + hpRowRec, 2) = ssToneType(j)
            Next
            hpRowRec = hpRowRec + UBound(ssToneType)
            Else
                For j = 1 To UBound(csToneType)
                    hpShtArr(j - 1 + hpRowRec, 2) = csToneType(j)
                Next
                hpRowRec = hpRowRec + UBound(csToneType)
        End If
    Next
    
    Dim excepCount As Integer '记录例外，特别指出现次数少的声调
    excepCount = 0
    Dim diffExcepArr() As String
    ReDim diffExcepArr(1 To 1)

    Dim srvyShtArr()
    ReDim srvyShtArr(1 To UBound(obArr) + 22, 1 To 50)
    Dim notExcep As Boolean, x As Integer, y As Integer, z As Integer
    For i = sylRowArr(1) To UBound(obArr)
        syl = obArr(i, sylCol)
        If syl <> "" And InStr(syl, " ") = 0 Then
            notExcep = True
            y = in_arr(initType, obArr(i, obSectCol + 1))
            x = in_arr(fnlType, obArr(i, obSectCol + 2))
            If x * y = 0 Then notExcep = False
            
            If in_arr(ssToneType, obArr(i, obSectCol + 3)) > 0 Then
                z = in_arr(ssToneType, obArr(i, obSectCol + 3))
                ElseIf in_arr(csToneType, obArr(i, obSectCol + 3)) > 0 Then
                    z = in_arr(csToneType, obArr(i, obSectCol + 3))
                Else
                    notExcep = False
            End If
            
            If notExcep Then '不是特殊情况才计入同音字表
                hpShtArr(hpFnlRowArr(x) + z - 1, y + 2) = hpShtArr(hpFnlRowArr(x) + z - 1, y + 2) + "//" + note_mod(obArr, i, Array(hzCol, wenbaiCol, shiyiCol, beizhuCol, obSectCol + 7))
                eaHpCount(hpFnlRowArr(x) + z - 1, y + 2) = eaHpCount(hpFnlRowArr(x) + z - 1, y + 2) + 1
                hpTbl(y, x, in_arr(toneType, obArr(i, obSectCol + 3))) = hpShtArr(hpFnlRowArr(x) + z - 1, y + 2)
                hpTblCnt(y, x, in_arr(toneType, obArr(i, obSectCol + 3))) = eaHpCount(hpFnlRowArr(x) + z - 1, y + 2)
                Else
                    notSame = True
                    For j = 1 To UBound(diffExcepArr)
                        If diffExcepArr(j) = obArr(i, sylCol) Then notSame = False
                    Next
                    If notSame Then
                        If diffExcepArr(UBound(diffExcepArr)) <> "" Then ReDim Preserve diffExcepArr(1 To UBound(diffExcepArr) + 1)
                        diffExcepArr(UBound(diffExcepArr)) = obArr(i, sylCol)
                    End If
                    
                    srvyShtArr(23 + excepCount, 1) = obArr(i, hzCol)
                    srvyShtArr(23 + excepCount, 2) = obArr(i, sylCol)
                    excepCount = excepCount + 1
            End If
        End If
    Next
    If excepCount > 0 Then
        srvyShtArr(21, 1) = "例外音节总数"
        srvyShtArr(21, 2) = excepCount
        srvyShtArr(22, 1) = "例外音节种数"
        srvyShtArr(22, 2) = UBound(diffExcepArr)
    End If
    
    Dim diffSylCount As Integer, tmprow As Integer, tmpcolumn As Integer
    diffSylCount = 0 '记录不重复的单音节的个数,下面的循环是为每个音节注上有多少个同音字
    For i = 2 To hpRowRec - 1
        For j = 3 To UBound(initType) + 2
            If eaHpCount(i, j) <> 0 Then
                hpShtArr(i, j) = CStr(eaHpCount(i, j)) + "/" + hpShtArr(i, j)
                diffSylCount = diffSylCount + 1
            End If
        Next
    Next '同音字表结束
    
    srvyShtArr(1, 1) = "包含例外"
    srvyShtArr(2, 1) = "记录的单音节"
    srvyShtArr(2, 2) = monSylCount
    srvyShtArr(3, 1) = "记录的多音节"
    srvyShtArr(3, 2) = polySylCount
    srvyShtArr(4, 1) = "拆多音节得到的单音节数"
    srvyShtArr(4, 2) = UBound(obArr) - obInitRowsCount
    srvyShtArr(5, 1) = ""
    srvyShtArr(5, 2) = ""
    srvyShtArr(6, 1) = "多音字/词个数"
    srvyShtArr(6, 2) = diffPpCount
    srvyShtArr(7, 1) = "音节数"
    srvyShtArr(7, 2) = UBound(sylRowArr)
    srvyShtArr(8, 1) = "声与韵的组合数"
    srvyShtArr(8, 2) = UBound(tonelessSylArr)
    srvyShtArr(9, 1) = "声母数"
    srvyShtArr(9, 2) = UBound(initArr)
    srvyShtArr(10, 1) = "韵母数"
    srvyShtArr(10, 2) = UBound(fnlArr)
    srvyShtArr(11, 1) = "声调数"
    srvyShtArr(11, 2) = UBound(toneArr)
    srvyShtArr(12, 1) = "辅音韵尾数"
    srvyShtArr(12, 2) = UBound(endArr)
    
    srvyShtArr(13, 1) = "不含例外"
    srvyShtArr(14, 1) = "同音字表音节数(不含例外)"
    srvyShtArr(14, 2) = diffSylCount
    srvyShtArr(15, 1) = "频数多于" + CStr(initExcepVal) + "的声母数"
    srvyShtArr(15, 2) = UBound(initType)
    srvyShtArr(16, 1) = "频数多于" + CStr(fnlExcepVal) + "的韵母数"
    srvyShtArr(16, 2) = UBound(fnlType)
    srvyShtArr(17, 1) = "频数多于" + CStr(toneExcepVal) + "的声调数"
    srvyShtArr(17, 2) = UBound(toneType)
    srvyShtArr(18, 1) = "最大可能音节数"
    maxPsbSylNum = UBound(initType) * (csFnlCount * csToneCount + ssFnlCount * ssToneCount)
    srvyShtArr(18, 2) = maxPsbSylNum
    srvyShtArr(19, 1) = "音节位利用率"
    srvyShtArr(19, 2) = diffSylCount / maxPsbSylNum
    
    srvyShtColTitle = Array("声母0", "总记录", "音节表", "声母", "韵母0", "总记录", "音节表", "韵母", "辅音韵尾", "总记录", "音节表", "声调0", "总记录", "音节表", "声调", "音节", "次数", "声韵组合", "次数", "声韵组合配调数", "次数")
    For j = 4 To 24
        srvyShtArr(1, j) = srvyShtColTitle(j - 4)
    Next
    
    '记录声、韵、调、尾在不重复的音节中出现的所有次数
    Dim posTmp As Integer
    Dim sylTblTonelessSylCountArr() As Integer
    ReDim sylTblTonelessSylCountArr(1 To UBound(tonelessSylArr))
    Dim sylTblInitCountArr() As Integer
    ReDim sylTblInitCountArr(1 To UBound(initArr))
    Dim sylTblFnlCountArr() As Integer
    ReDim sylTblFnlCountArr(1 To UBound(fnlArr))
    Dim sylTblToneCountArr() As Integer
    ReDim sylTblToneCountArr(1 To UBound(toneArr))
    Dim sylTblEndCountArr() As Integer
    ReDim sylTblEndCountArr(1 To UBound(endArr))
    For i = 1 To UBound(sylRowArr)
        posTmp = in_arr(tonelessSylArr, obArr(sylRowArr(i), obSectCol))
        If posTmp > 0 Then sylTblTonelessSylCountArr(posTmp) = sylTblTonelessSylCountArr(posTmp) + 1
        posTmp = in_arr(initArr, obArr(sylRowArr(i), obSectCol + 1))
        If posTmp > 0 Then sylTblInitCountArr(posTmp) = sylTblInitCountArr(posTmp) + 1
        posTmp = in_arr(fnlArr, obArr(sylRowArr(i), obSectCol + 2))
        If posTmp > 0 Then sylTblFnlCountArr(posTmp) = sylTblFnlCountArr(posTmp) + 1
        posTmp = in_arr(toneArr, obArr(sylRowArr(i), obSectCol + 3))
        If posTmp > 0 Then sylTblToneCountArr(posTmp) = sylTblToneCountArr(posTmp) + 1
        posTmp = in_arr(endArr, obArr(sylRowArr(i), obSectCol + 4))
        If posTmp > 0 Then sylTblEndCountArr(posTmp) = sylTblEndCountArr(posTmp) + 1
    Next
    
    For i = 1 To UBound(initArr)
        srvyShtArr(i + 1, 4) = initArr(i)
        srvyShtArr(i + 1, 5) = initCountArr(i)
        srvyShtArr(i + 1, 6) = sylTblInitCountArr(i)
        If i <= UBound(initType) Then srvyShtArr(i + 1, 7) = initArr(i)
    Next
    For i = 1 To UBound(fnlArr)
        srvyShtArr(i + 1, 8) = fnlArr(i)
        srvyShtArr(i + 1, 9) = fnlCountArr(i)
        srvyShtArr(i + 1, 10) = sylTblFnlCountArr(i)
        If i <= UBound(fnlType) Then srvyShtArr(i + 1, 11) = fnlArr(i)
    Next
    For i = 1 To UBound(endArr)
        srvyShtArr(i + 1, 12) = endArr(i)
        srvyShtArr(i + 1, 13) = endCountArr(i)
        srvyShtArr(i + 1, 14) = sylTblEndCountArr(i)
    Next
    
    For i = 1 To UBound(toneArr)
        srvyShtArr(i + 1, 15) = toneArr(i)
        srvyShtArr(i + 1, 16) = toneCountArr(i)
        srvyShtArr(i + 1, 17) = sylTblToneCountArr(i)
        If i <= UBound(toneType) Then srvyShtArr(i + 1, 18) = toneArr(i)
    Next
    
    countArrTmp = sylCountArr
    sylSortedPos = range_arr(1, UBound(sylRowArr), 1)
    Call quick_sort(countArrTmp, sylSortedPos, 1, UBound(countArrTmp))
    
    For i = 1 To UBound(sylRowArr)
        srvyShtArr(i + 1, 19) = obArr(sylRowArr(sylSortedPos(i)), obSectCol) + obArr(sylRowArr(sylSortedPos(i)), obSectCol + 3)
        srvyShtArr(i + 1, 20) = sylCountArr(sylSortedPos(i))
    Next
    For i = 1 To UBound(tonelessSylArr)
        srvyShtArr(i + 1, 21) = tonelessSylArr(i)
        srvyShtArr(i + 1, 22) = tonelessSylCountArr(i)
    Next
    Call quick_sort(sylTblTonelessSylCountArr, tonelessSylArr, 1, UBound(sylTblTonelessSylCountArr))
    For i = 1 To UBound(tonelessSylArr)
        srvyShtArr(i + 1, 23) = tonelessSylArr(i)
        srvyShtArr(i + 1, 24) = sylTblTonelessSylCountArr(i)
    Next '概况
    
    Dim toneShtArr() As String '调型分布,不计入例外
    ReDim toneShtArr(1 To UBound(obArr) + UBound(fnlType) + 30, 1 To 256)
    For i = 1 To UBound(sylRowArr)
        y = in_arr(initType, obArr(sylRowArr(i), obSectCol + 1))
        x = in_arr(fnlType, obArr(sylRowArr(i), obSectCol + 2))
        If x * y <> 0 Then toneShtArr(x + 1, y + 1) = toneShtArr(x + 1, y + 1) + "/" + obArr(sylRowArr(i), obSectCol + 3) '调型分布
    Next
    For i = 1 To UBound(initType)
        toneShtArr(1, i + 1) = initType(i)
    Next
    For i = 1 To UBound(fnlType)
        toneShtArr(i + 1, 1) = fnlType(i)
    Next '调型分布
    
    Call table_example(4, initSysTbl, initType, fnlType, toneType, hpTbl, hpTblCnt, 1)
    Call table_example(4, fnlSysTbl, initType, fnlType, toneType, hpTbl, hpTblCnt, 2)
    toneSysTbl = toneType
    Call tone_example_table(5, ByVal csToneCount, toneSysTbl, initType, fnlType, hpTbl)
    Dim phonSysShtArr() As String '音系表格
    ReDim phonSysShtArr(1 To UBound(initSysTbl, 1) + UBound(fnlSysTbl, 1) + UBound(toneSysTbl) + 3, 1 To UBound(initSysTbl, 2))
    phonSysShtArr(1, 1) = "声母(" + CStr(UBound(initType)) + "个)"
    Call fill_arr(phonSysShtArr, initSysTbl, 2, 1, 2)
    phonSysShtArr(UBound(initSysTbl, 1) + 2, 1) = "韵母(" + CStr(UBound(fnlType)) + "个)"
    Call fill_arr(phonSysShtArr, fnlSysTbl, UBound(initSysTbl, 1) + 3, 1, 2)
    phonSysShtArr(UBound(initSysTbl, 1) + UBound(fnlSysTbl, 1) + 3, 1) = "声调(" + CStr(UBound(toneSysTbl)) + "个)"
    Call fill_arr(phonSysShtArr, toneSysTbl, UBound(initSysTbl, 1) + UBound(fnlSysTbl, 1) + 4, 1, 11)
    
    If polySylCount > 0 Then '在音系出来后才能做调型组合.同时判断有没有多音节
        Dim biSylCount As Integer, triSylCount As Integer, quadSylCount As Integer
        Dim biSyl()
        ReDim biSyl(1 To UBound(polySylArr), 3)
        Dim triSyl()
        ReDim triSyl(1 To UBound(polySylArr), 4)
        Dim quadSylType As String
        '定位点
        biSylCoord = Array(UBound(fnlType) + 5, 1)
        triSylCoord = Array(UBound(fnlType) + 5, UBound(toneType) + 4)
        quadSylCoord = Array(UBound(fnlType) + 5, 2 * UBound(toneType) + 7)
        For i = 1 To UBound(polySylArr)
            syl = obArr(polySylArr(i), sylCol)
            If syl Like "* * * *" Then
                s4 = Split(syl, " ")
                quadSylType = pick_tone(s4(0)) + " " + pick_tone(s4(1)) + " " + pick_tone(s4(2)) + " " + pick_tone(s4(3))
                toneShtArr(quadSylCoord(0) + 2 + quadSylCount, quadSylCoord(1) + 1) = quadSylType
                toneShtArr(quadSylCoord(0) + 2 + quadSylCount, quadSylCoord(1) + 2) = obArr(polySylArr(i), hzCol)
                toneShtArr(quadSylCoord(0) + 2 + quadSylCount, quadSylCoord(1) + 3) = syl
                quadSylCount = quadSylCount + 1
                ElseIf syl Like "* * *" Then
                    triSylCount = triSylCount + 1
                    s3 = Split(syl, " ")
                    triSyl(triSylCount, 0) = syl
                    triSyl(triSylCount, 1) = pick_tone(s3(0))
                    triSyl(triSylCount, 2) = pick_tone(s3(1))
                    triSyl(triSylCount, 3) = pick_tone(s3(2))
                    For j = 1 To 3 '如果含有特殊的声调，就要区分开，不列入声调组合表
                        If in_arr(toneType, triSyl(triSylCount, j)) = 0 Then
                            triSyl(triSylCount, 1) = triSyl(triSylCount, 1) + "特殊"
                            Exit For
                        End If
                    Next
                    triSyl(triSylCount, 4) = obArr(polySylArr(i), hzCol)
                    Else
                        biSylCount = biSylCount + 1
                        s2 = Split(syl, " ")
                        biSyl(biSylCount, 0) = syl
                        biSyl(biSylCount, 1) = pick_tone(s2(0))
                        biSyl(biSylCount, 2) = pick_tone(s2(1))
                        For j = 1 To 2 '如果含有特殊的声调，就要区分开，不列入声调组合表
                            If in_arr(toneType, biSyl(biSylCount, j)) = 0 Then
                                biSyl(biSylCount, 1) = biSyl(biSylCount, 1) + "特殊"
                                Exit For
                            End If
                        Next
                        biSyl(biSylCount, 3) = obArr(polySylArr(i), hzCol)
            End If
        Next
    
        toneShtArr(quadSylCoord(0), quadSylCoord(1)) = quadSylCount
        toneShtArr(quadSylCoord(0) - 1, quadSylCoord(1)) = "四音节"
        toneShtArr(quadSylCoord(0) - 1, quadSylCoord(1) + 1) = "包含特殊的总数目"
        toneShtArr(quadSylCoord(0) - 1, quadSylCoord(1) + 2) = quadSylCount
    
        '双音节
        toneShtArr(biSylCoord(0) + 1, biSylCoord(1)) = "第一音节"
        toneShtArr(biSylCoord(0), biSylCoord(1) + 1) = "第二音节"
        Dim h As Integer, biSylTypeCount As Integer, triSylTypeCount As Integer
        Dim hi() As Integer     'h用来统计第一个声调相同时第二个声调的最大峰，第二音节每种声调都有一个对应的hi记录位置
        ReDim hi(UBound(toneType))
        For n = 0 To UBound(toneType)
            hi(n) = biSylCoord(0) + 2
        Next
        For j = 1 To UBound(toneType)
            toneShtArr(biSylCoord(0) + 1, biSylCoord(1) + j) = toneType(j) '第二音节的声调
            h = Application.WorksheetFunction.Max(hi)
            toneShtArr(h, biSylCoord(1)) = toneType(j) '第一音节的声调
            For n = 0 To UBound(toneType)
                hi(n) = h
            Next
            For i = 1 To biSylCount
                If biSyl(i, 1) = toneType(j) Then
                    x = Application.Match(biSyl(i, 2), toneType, 0)
                    toneShtArr(hi(x), x + biSylCoord(1)) = biSyl(i, 0)
                    toneShtArr(hi(x) + 1, x + biSylCoord(1)) = biSyl(i, 3)
                    hi(x) = hi(x) + 2
                End If
            Next
            For i = 1 To UBound(toneType) - 1 '记录这个调型有没有
                If hi(i) > h Then biSylTypeCount = biSylTypeCount + 1
            Next
        Next
        toneShtArr(biSylCoord(0), biSylCoord(1)) = CStr(biSylTypeCount) + "种"
        toneShtArr(biSylCoord(0) - 1, biSylCoord(1)) = "双音节"
        toneShtArr(biSylCoord(0) - 1, biSylCoord(1) + 1) = "包含特殊的总数目"
        toneShtArr(biSylCoord(0) - 1, biSylCoord(1) + 2) = biSylCount
    
        '三音节
        toneShtArr(triSylCoord(0) + 1, triSylCoord(1)) = "第一音节"
        toneShtArr(triSylCoord(0), triSylCoord(1) + 1) = "第二音节"
        toneShtArr(triSylCoord(0) + 1, triSylCoord(1) + UBound(toneType)) = "第三音节"
        Dim patternNotExist As Boolean
        For n = 0 To UBound(toneType)
            hi(n) = triSylCoord(0) + 2
        Next
        For k = 1 To UBound(toneType)
            toneShtArr(triSylCoord(0) + 1, triSylCoord(1) + k) = toneType(k) '第二音节的声调
            toneShtArr(Application.WorksheetFunction.Max(hi), triSylCoord(1)) = toneType(k) '第一音节的声调
            For j = 1 To UBound(toneType)
                h = Application.WorksheetFunction.Max(hi)
                toneShtArr(h, triSylCoord(1) + UBound(toneType)) = toneType(j) '第三音节的声调
                For n = 0 To UBound(toneType)
                    hi(n) = h
                Next
                For i = 1 To triSylCount
                    If triSyl(i, 1) = toneType(k) And triSyl(i, 3) = toneType(j) Then
                        x = Application.Match(triSyl(i, 2), toneType, 0)
                        toneShtArr(hi(x), x + triSylCoord(1)) = triSyl(i, 0)
                        toneShtArr(hi(x) + 1, x + triSylCoord(1)) = triSyl(i, 4)
                        hi(x) = hi(x) + 2
                    End If
                Next
            
                patternNotExist = True
                For i = 1 To UBound(toneType) '记录这个调型有没有,如果没有就空出空格
                    If hi(i) > h Then
                        triSylTypeCount = triSylTypeCount + 1
                        judge = False
                    End If
                Next
                If patternNotExist Then hi(0) = h + 2
            Next
        Next
        toneShtArr(triSylCoord(0), triSylCoord(1)) = CStr(triSylTypeCount) + "种"
        toneShtArr(triSylCoord(0) - 1, triSylCoord(1)) = "三音节"
        toneShtArr(triSylCoord(0) - 1, triSylCoord(1) + 1) = "包含特殊的总数目"
        toneShtArr(triSylCoord(0) - 1, triSylCoord(1) + 2) = triSylCount
    End If
    
    '生成表格
    Dim obSht As String, srvySht As String, hpSht As String, toneSht As String, phonSysSht As String
    obSht = "切分后"
    Sheets.Add.Name = obSht
    srvySht = "概况"
    hpSht = "同音字表"
    toneSht = "调型"
    phonSysSht = "音系"
	sheets(tgSht).select
    Sheets.Add.Name = srvySht
    Sheets.Add.Name = hpSht
    Sheets.Add.Name = toneSht
    Sheets.Add.Name = phonSysSht
    
    Sheets(obSht).[a1].Resize(UBound(obArr, 1), UBound(obArr, 2)) = obArr
    Sheets(srvySht).[a1].Resize(UBound(srvyShtArr, 1), UBound(srvyShtArr, 2)) = srvyShtArr
    Sheets(hpSht).[a1].Resize(UBound(hpShtArr, 1), UBound(hpShtArr, 2)) = hpShtArr
    Sheets(toneSht).[a1].Resize(UBound(toneShtArr, 1), UBound(toneShtArr, 2)) = toneShtArr
    Sheets(phonSysSht).[a1].Resize(UBound(phonSysShtArr, 1), UBound(phonSysShtArr, 2)) = phonSysShtArr
    
    MsgBox "总运行时间为" & Timer - beginMoment
End Sub

Public Function std_syl(ByVal syl As String) As String
    Dim m As Integer, n As Integer, sensor1 As Integer, sensor2 As Integer, letter As String
    For i = 1 To Len(syl)
        letter = Mid(syl, i, 1)
        If InStr("qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM0123456789~!?@#$^&<>*=+", letter) = 0 Then syl = Replace(syl, letter, "")
    Next
    sensor2 = 0
    i = 1
    Do While i <= Len(syl)
        sensor1 = sensor2 '判断前一个字符是否是数字
        sensor2 = 0
        If InStr("0123456789", Mid(syl, i, 1)) > 0 Then sensor2 = 1 '当前是数字
        If sensor1 = 1 And sensor2 = 0 Then
            syl = left(syl, i - 1) + " " + right(syl, Len(syl) - i + 1)
            i = i + 1
        End If
        i = i + 1
    Loop
    If InStr(syl, " ?") > 0 Then syl = Replace(syl, " ?", "")
    If InStr(syl, " !") > 0 Then syl = Replace(syl, " !", "")
    If InStr(syl, "! ") > 0 Then syl = Replace(syl, "! ", "")
    std_syl = Trim(syl)
End Function

Public Function syl_sect(ByVal syl As String) '将音节拆成CVCT格式
    Dim sylSectArr(1 To 4) As String, sectStr As String, i As Integer, j As Integer
    sectStr = syl
    sectStr = Replace(sectStr, "ng", "三")
    sectStr = Replace(sectStr, "w$", "一")
    sectStr = Replace(sectStr, "y$", "二")
    i = Len(sectStr)
    Do While InStr("0123456789", Mid(sectStr, i, 1)) > 0
        i = i - 1
        If i = 0 Then
            syl_sect = Array("有误", "有误", "有误", sectStr, "有误")
            Exit Function
        End If
    Loop
    sectStr = left(sectStr, i) + "," + right(sectStr, Len(sectStr) - i) '调为空的情况也可以算出来
    If InStr("ptk?", Mid(sectStr, i, 1)) > 0 Then
        sectStr = sectStr + "入"
        sectStr = left(sectStr, i - 1) + "," + right(sectStr, Len(sectStr) - i + 1)
        ElseIf InStr("mn三", Mid(sectStr, i, 1)) > 0 Then
            sectStr = left(sectStr, i - 1) + "," + right(sectStr, Len(sectStr) - i + 1)
        Else
            sectStr = left(sectStr, i) + "," + right(sectStr, Len(sectStr) - i)
    End If
    j = 1
    Do While j < i And InStr("iyuwIYUeoEaA", Mid(sectStr, j, 1)) = 0 '如果辅音韵尾前的逗号是寻找声母的上限位置
        j = j + 1
    Loop
    sectStr = left(sectStr, j - 1) + "," + right(sectStr, Len(sectStr) - j + 1) '这个逗号分声母
    sectStr = Replace(sectStr, "三", "ng")
    sectStr = Replace(sectStr, "一", "w$")
    sectStr = Replace(sectStr, "二", "y$")
    arr = Split(sectStr, ",")
    conFnlArr = Split("m,n,ng,l,z,v", ",")
    If arr(0) = "" Then arr(0) = "?"
    If arr(1) = "" Then '声化韵的情况
        If in_arr(conFnlArr, arr(2)) > 0 Then
            arr(1) = arr(2)
            arr(2) = ""
            Else
                syl_sect = Array("有误", arr(0), "有误", arr(3), arr(2))
                Exit Function
        End If
    End If
    syl_sect = Array(arr(0) + arr(1) + arr(2), arr(0), arr(1) + arr(2), arr(3), arr(2))
End Function

Public Function pick_tone(ByVal syl As Variant) As String
    Dim j As Integer, tmpstr As String, tone As String
    j = Len(syl)
    tmpstr = right(syl, 1)
    Do While InStr("0123456789", tmpstr) > 0
        tone = tmpstr + tone '如果是数字就加到dstr里面
        j = j - 1
        tmpstr = Mid(syl, j, 1) '取yin的每一个字符，并且是倒序取，会快一些
    Loop
    If tone = "" Then tone = "空"
    If InStr("?ptk", tmpstr) > 0 Then tone = tone + "入"
    pick_tone = tone
End Function

Public Function bubble_sort(ByRef arr, ByRef lnkArr)
    Dim i As Integer, j As Integer, swapLnk As Boolean
    For i = 1 To UBound(arr)
        For j = 1 To UBound(arr) - i
            If arr(j) < arr(j + 1) Then
                Call swap(arr, j, j + 1)
                Call swap(lnkArr, j, j + 1)
            End If
        Next
    Next
End Function

Public Function in_arr(ByRef arr, ByVal tg) As Integer '返回第一次出现时的位置
    Dim i As Integer, loc As Integer, tail As Integer, start As Integer
    loc = 0
    On Error Resume Next
    start = LBound(arr)
    If Err.Number = 0 Then
        tail = UBound(arr)
        For i = 0 To tail - start
            If tg = arr(tail - i) Then
                loc = tail - i - start + 1
                Exit For
            End If
        Next
        Else
            loc = -1
    End If
    in_arr = loc
End Function

Public Function note_mod(ByRef arr, ByVal k As Integer, ByRef idx As Variant) As String
    Dim i As Integer
    For i = LBound(idx) To UBound(idx)
        If Val(idx(i)) >= LBound(arr, 2) And Val(idx(i)) <= UBound(arr, 2) Then note_mod = note_mod + " " + CStr(arr(k, Val(idx(i))))
    Next
End Function

Public Function catg_count(ByRef itemArr, ByRef countArr, ByVal cmpStr As String)
    Dim notSame As Boolean, j As Integer
    notSame = True
    j = in_arr(itemArr, cmpStr)
    If j > 0 Then
        notSame = False
        countArr(j) = countArr(j) + 1
    End If
    If notSame Then
        ReDim Preserve itemArr(1 To UBound(itemArr) + 1)
        ReDim Preserve countArr(1 To UBound(countArr) + 1)
        itemArr(UBound(itemArr)) = cmpStr
        countArr(UBound(countArr)) = 1
    End If
End Function

Public Function find_type(ByRef typeArr, ByRef countArr() As Integer, ByRef arr, ByVal excepVal As Integer)
    Dim i As Integer
    ReDim typeArr(1 To 1)
    typeArr(1) = arr(1)
    For i = 2 To UBound(arr)
        If countArr(i) > excepVal Then
            ReDim Preserve typeArr(1 To UBound(typeArr) + 1)
            typeArr(UBound(typeArr)) = arr(i)
            Else
                Exit For
        End If
    Next
End Function

'声、韵、调按照语音学排序
Public Function init_phon_sort(ByRef typeArr)
    Dim i As Integer, j As Integer, rk As Integer
    initTbl = init_tbl(typeArr)
    For i = 1 To UBound(initTbl, 1)
        For j = 1 To UBound(initTbl, 2)
            If initTbl(i, j) <> "" Then
                rk = rk + 1
                typeArr(rk) = initTbl(i, j)
            End If
        Next
    Next
End Function

Public Function init_tbl(ByRef typeArr)
    Dim i As Integer, j As Integer
    ipaInitStr = "p,ph,b,bh,,,,,m,,,,,,p*,b*,,;,,,,pf,pfh,bv,bvh,mg,,,,,,f,v,v$,;,,,,t>,t>h,d>,d>h,,,,,,,s>,z>,,;,,,,ts,tsh,dz,dzh,,,,,,,s,z,,;t,th,d,dh,,,,,n,r,r*,l,ls,l#,,,r$,;tr,trh,dr,drh,tsr,tsrh,dzr,dzrh,nr,,r^,lr,,,sr,zr,rr,;,,,,tss,tssh,dzz,dzzh,,,,,,,ss,zz,,;tj,tjh,dj,djh,tcj,tcjh,dzj,dzjh,nj,,,,,,cj,zj,,;c,ch,c!,c!h,,,,,nc,,,lc,,,c#,jj,j,y$;k,kh,g,gh,,,,,ng,,,,,,x,x!,w!,w$;q,qh,G,Gh,,,,,N,R,,,,,X,X!,,;,,,,,,,,,,,,,,h*,h*!,,;?,?h,,,,,,,,,,,,,h,h!,,"
    arr1 = Split(ipaInitStr, ";")
    Dim ipaInitArr(1 To 13, 1 To 18) As String
    For i = 0 To 12
        arr2 = Split(arr1(i), ",")
        For j = 0 To 17
            If in_arr(typeArr, arr2(j)) > 0 Then
                ipaInitArr(i + 1, j + 1) = arr2(j) '记录了的声母才点亮
            End If
        Next
    Next
    Dim rowExist() As Integer, colExist() As Integer
    ReDim rowExist(1 To 1)
    ReDim colExist(1 To 1)
    For i = 1 To 13
        NotEmpty = False
        For j = 1 To 18
            If ipaInitArr(i, j) <> "" Then
                NotEmpty = True
                Exit For
            End If
        Next
        If NotEmpty Then
            If rowExist(UBound(rowExist)) <> 0 Then ReDim Preserve rowExist(1 To UBound(rowExist) + 1)
            rowExist(UBound(rowExist)) = i
        End If
    Next
    For j = 1 To 18
        NotEmpty = False
        For i = 1 To 13
            If ipaInitArr(i, j) <> "" Then
                NotEmpty = True
                Exit For
            End If
        Next
        If NotEmpty Then
            If colExist(UBound(colExist)) <> 0 Then ReDim Preserve colExist(1 To UBound(colExist) + 1)
            colExist(UBound(colExist)) = j
        End If
    Next
    Dim initTbl() As String
    ReDim initTbl(1 To UBound(rowExist), 1 To UBound(colExist))
    For i = 1 To UBound(rowExist)
        For j = 1 To UBound(colExist)
            initTbl(i, j) = ipaInitArr(rowExist(i), colExist(j))
        Next
    Next
    init_tbl = initTbl
End Function

Public Function fnl_phon_sort(ByRef typeArr)
    Dim i As Integer, j As Integer
    sortedFnlArr = typeArr
    For i = 1 To UBound(typeArr)
        For j = 1 To UBound(typeArr) - i
            If fnl_num(typeArr(j)) > fnl_num(typeArr(j + 1)) Then Call swap(typeArr, j, j + 1)
        Next
    Next
End Function

Public Function fnl_tbl(ByRef sortedFnlArr)
    Dim fnlTbl(), numRec As Long, rowCnt As Integer, csCount As Integer, fnlNum As Long
    For i = 1 To UBound(sortedFnlArr)
        fnlNum = fnl_num(sortedFnlArr(i))
        If fnlNum > numRec Then
            rowCnt = rowCnt + 1
            ReDim Preserve fnlTbl(1 To 4, 1 To rowCnt)
            numRec = fnlNum + 5
        End If
        fnlTbl(fnlNum Mod 10, rowCnt) = sortedFnlArr(i)
        If fnlNum \ 100000 >= 4 And fnlNum \ 100000 <= 7 Then csCount = csCount + 1 '计算促声韵的数目
    Next
    fnlTbl = array_transpose(fnlTbl)
    fnl_tbl = Array(fnlTbl, csCount)
End Function

Public Function fnl_num(ByVal fnl As String) As Long
    sectArr = fnl_sect(fnl)
    fnl_num = in_arr(Split("i,u,y", ","), sectArr(1)) + 1 + vowel_num(sectArr(2)) * 10 + vowel_num(sectArr(3)) * 100 + in_arr(Split("m,n,ng,p,t,k,?", ","), sectArr(4)) * 100000
    If fnl_num = 0 Then fnl_num = Int(91 * Rnd + 10) * 1000000
End Function

Public Function fnl_sect(ByVal fnl As String) '将韵母拆成vvvc格式
    Dim sectArr(1 To 4) As String, vvv As String, letter As String, count As Integer
    If in_arr(Split("z,v,l,m,n,ng", ","), fnl) > 0 Then '声化韵
        sectArr(2) = fnl
        fnl_sect = sectArr
        Exit Function
    End If
    fnl = Replace(fnl, "ng", "三")
    If in_arr(Split("m,n,三,p,t,k,?", ","), right(fnl, 1)) > 0 Then
        sectArr(4) = Replace(right(fnl, 1), "三", "ng")
        vvv = left(fnl, Len(fnl) - 1)
        Else
            vvv = fnl
    End If
    i = 1
    Do While i <= Len(vvv)
        letter = Mid(vvv, i, 1) 'mid函数起始即从1开始
        If InStr("iyuwIYUeoEaA", letter) > 0 Then
            vvv = left(vvv, i - 1) + "," + right(vvv, Len(vvv) - i + 1)
            i = i + 1
        End If
        i = i + 1
    Loop
    vvv = right(vvv, Len(vvv) - 1)
    If vvv <> "" Then
        arr = Split(vvv, ",")
        If in_arr(Split("i,u,y", ","), arr(0)) > 0 Then
            sectArr(1) = arr(0)
            count = 1
            Else
                sectArr(2) = arr(0)
                count = 2
        End If
        If UBound(arr) >= 1 Then
            For i = 1 To UBound(arr)
                sectArr(i + count) = arr(i)
            Next
        End If
    End If
    fnl_sect = sectArr
End Function

Public Function vowel_num(ByVal vowel As String) As Long
    If vowel = "" Then Exit Function
    Dim pos As Integer, ipaFnlStr As String
    pos = in_arr(Split("z,v,l,m,n,ng", ","), vowel)
    If pos > 0 Then
        vowel_num = pos * 1000000
        Exit Function
    End If
    ipaFnlStr = "i~,y~,i#~,u#~,u=~,u~;I~,Y~,,,U~,;e~,e@~,e#~,o#~,e>~,o~;E~,,e=~,,,o=~;e+~,e+@~,e+#~,o+#~,o+$~,o+~;a^~,,A^~,,,;a~,a@~,A~,,a>~,a>@~;i<~,i>~,y<~,y>~,,;i,y,i#,u#,u=,u;I,Y,,,U,;e,e@,e#,o#,e>,o;E,,e=,,,o=;e+,e+@,e+#,o+#,o+$,o+;a^,,A^,,,;a,a@,A,,a>,a>@;i<,i>,y<,y>,,"
    arr1 = Split(ipaFnlStr, ";")
    For i = 0 To UBound(arr1)
        pos = in_arr(Split(arr1(i), ","), vowel)
        If pos > 0 Then
            vowel_num = (UBound(arr1) - i) * 10 + pos
            If InStr(vowel, "~") > 0 Then vowel_num = vowel_num * 100
            Exit For
        End If
    Next
End Function

Public Function tone_phon_sort(ByRef typeArr)
    Dim i As Integer, j As Integer
    For i = 1 To UBound(typeArr)
        For j = 1 To UBound(typeArr) - i
            If tone_num(typeArr(j)) < tone_num(typeArr(j + 1)) Then Call swap(typeArr, j, j + 1)
        Next
    Next
End Function

Public Function tone_num(ByVal tone As String)  '平、升、降的顺序，曲折调按照前一段判断平升降，短调按降调处理
    Dim k As Single, trend As Integer, ss As Integer
    ss = 10000
    If right(tone, 1) = "入" Then
        tone = left(tone, Len(tone) - 1)
        ss = 0
    End If
    tone = tone + "0"
    If Len(tone) > 1 Then trend = left(tone, 1) - Mid(tone, 2, 1)
    If trend > 0 Then
        k = 0.000001
        ElseIf trend < 0 Then
            k = 0.1
        Else
            k = 10000
    End If
    tone_num = k * Val(tone) + ss
End Function

Function quick_sort(ByRef arr, ByRef lnkArr, ByVal left As Integer, ByVal right As Integer)
    If left >= right Then Exit Function
    Dim pivotIdx As Integer
    pivotIdx = partition(arr, lnkArr, left, right)
    Call quick_sort(arr, lnkArr, left, pivotIdx - 1)
    Call quick_sort(arr, lnkArr, pivotIdx + 1, right)
End Function

Function partition(ByRef arr, ByRef lnkArr, ByVal left As Integer, ByVal right As Integer)
    Dim i As Integer, tail As Integer
    pivot = arr(right)
    tail = left - 1
    For i = left To right - 1
        If arr(i) >= pivot Then
            tail = tail + 1 'tail记录有多少个大于pivot的数，然后每次都把小的数换到左边数组的最右边
            If tail <> i Then
                Call swap(arr, tail, i)
                Call swap(lnkArr, tail, i)
            End If
        End If
    Next
    Call swap(arr, tail + 1, right) '把pivot也换到左边数组的最右边
    Call swap(lnkArr, tail + 1, right)
    partition = tail + 1
End Function

Function swap(ByRef arr, ByVal i As Integer, ByVal j As Integer)
    tmp = arr(i)
    arr(i) = arr(j)
    arr(j) = tmp
End Function

Function range_arr(ByVal headPos As Long, ByVal length As Long, ByVal startNum As Long, Optional ByVal interval As Integer = 1)
    Dim arr()
    ReDim arr(headPos To length + headPos - 1)
    For i = headPos To length - 1 + headPos
        arr(i) = startNum + interval * (i - headPos)
    Next
    range_arr = arr
End Function

Function array_transpose(ByRef arr) '只转置二维数组，没有vba自带的transpose的限制
    Dim i As Integer, j As Integer, bidi As Integer
    Dim tArr()
    If IsArray(arr) Then
        On Error Resume Next
        bidi = LBound(arr, 2)
        If Err.Number = 0 Then
            ReDim tArr(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    tArr(j, i) = arr(i, j)
                Next
            Next
            array_transpose = tArr
        End If
    End If
End Function

Public Function fill_arr(ByRef arr, ByRef fillerArr, ByVal row As Integer, ByVal col As Integer, Optional pattern As Integer = 2) '默认模式是二维填二维
    Dim i As Integer, j As Integer
    If IsArray(fillerArr) Then
        If pattern = 2 Then
            If LBound(arr, 1) <= row And LBound(arr, 2) <= col And UBound(arr, 1) >= row + UBound(fillerArr, 1) - LBound(fillerArr, 1) And UBound(arr, 2) >= col + UBound(fillerArr, 2) - LBound(fillerArr, 2) Then
                For i = LBound(fillerArr, 1) To UBound(fillerArr, 1)
                    For j = LBound(fillerArr, 2) To UBound(fillerArr, 2)
                        arr(row + i - LBound(fillerArr, 1), col + j - LBound(fillerArr, 2)) = fillerArr(i, j)
                    Next
                Next
            End If
            ElseIf pattern = 1 Then
                If LBound(arr) <= row And UBound(arr) >= row + UBound(fillerArr) - LBound(fillerArr) Then
                    For i = LBound(fillerArr) To UBound(fillerArr)
                        arr(row + i - LBound(fillerArr)) = fillerArr(i)
                    Next
                End If
            ElseIf pattern = 11 Then
                If LBound(arr, 1) <= row And UBound(arr, 1) >= row + UBound(fillerArr) - LBound(fillerArr) Then
                For i = LBound(fillerArr) To UBound(fillerArr)
                        arr(row + i - LBound(fillerArr), col) = fillerArr(i)
                Next
                End If
            Else
                If LBound(arr, 2) <= col And UBound(arr, 2) >= col + UBound(fillerArr) - LBound(fillerArr) Then
                    For i = LBound(fillerArr) To UBound(fillerArr)
                        arr(row, col + i - LBound(fillerArr)) = fillerArr(i)
                    Next
                End If
        End If
    End If
End Function

Public Function table_example(ByVal cake As Integer, ByRef tblArr, ByRef initType, ByRef fnlType, ByRef toneType, ByRef hpTbl, ByRef hpTblCnt, Optional mode As Integer = 1)
    Dim i As Integer, j As Integer
    If IsArray(tblArr) Then
        For i = LBound(tblArr, 1) To UBound(tblArr, 1)
            For j = LBound(tblArr, 2) To UBound(tblArr, 2)
                If tblArr(i, j) <> "" Then
                    If mode = 1 Then Call init_plus_example(cake, tblArr, i, j, initType, fnlType, toneType, hpTbl, hpTblCnt)
                    If mode = 2 Then Call fnl_plus_example(cake, tblArr, i, j, initType, fnlType, toneType, hpTbl, hpTblCnt)
                End If
            Next
        Next
    End If
End Function

Public Function init_plus_example(ByVal cake As Integer, ByRef tblArr, ByVal row As Integer, ByVal col As Integer, ByRef initType, ByRef fnlType, ByRef toneType, ByRef hpTbl, ByRef hpTblCnt) As String '为单个声母找例字
    Dim countArr() As Integer, pickArr() As String
    ReDim countArr(LBound(fnlType) To UBound(fnlType))
    ReDim pickArr(LBound(fnlType) To UBound(fnlType))
    Dim x As Integer, y As Integer, z As Integer, i As Integer, k As Integer, eg As Integer, timesRec As Integer
    x = in_arr(initType, tblArr(row, col))
    For y = LBound(fnlType) To UBound(fnlType)
        For z = LBound(toneType) To UBound(toneType)
            countArr(y) = countArr(y) + hpTblCnt(x, y, z)
            pickArr(y) = pickArr(y) + hpTbl(x, y, z)
        Next
    Next
    sortedPos = range_arr(1, UBound(fnlType), 1)
    Call quick_sort(countArr, sortedPos, LBound(fnlType), UBound(fnlType))
    exampleNum = cut_cake(cake, countArr)
    For i = LBound(exampleNum) To UBound(exampleNum)
        selectArr = Split(pickArr(sortedPos(i)), "//")
        For k = 1 To exampleNum(i)
            eg = Int((UBound(selectArr) - 1) * Rnd + 1)
            timesRec = 1
            Do While in_arr(selectArr(eg), "pp") > 0 Or Len(selectArr(eg)) >= 15
                If timesRec > 10 Then Exit Do
                eg = Int((UBound(selectArr) - 1) * Rnd + 1)
                timesRec = timesRec + 1
            Loop
            tblArr(row, col) = tblArr(row, col) + " //" + selectArr(eg)
        Next
    Next
End Function

Public Function fnl_plus_example(ByVal cake As Integer, ByRef tblArr, ByVal row As Integer, ByVal col As Integer, ByRef initType, ByRef fnlType, ByRef toneType, ByRef hpTbl, ByRef hpTblCnt) As String '为单个声母找例字
    Dim countArr() As Integer, pickArr() As String
    ReDim countArr(LBound(initType) To UBound(initType))
    ReDim pickArr(LBound(initType) To UBound(initType))
    Dim x As Integer, y As Integer, z As Integer, i As Integer, k As Integer, eg As Integer, timesRec As Integer
    y = in_arr(fnlType, tblArr(row, col))
    For x = LBound(initType) To UBound(initType)
        For z = LBound(toneType) To UBound(toneType)
            countArr(x) = countArr(x) + hpTblCnt(x, y, z)
            pickArr(x) = pickArr(x) + hpTbl(x, y, z)
        Next
    Next
    sortedPos = range_arr(1, UBound(initType), 1)
    Call quick_sort(countArr, sortedPos, LBound(initType), UBound(initType))
    exampleNum = cut_cake(cake, countArr)
    For i = LBound(exampleNum) To UBound(exampleNum)
        selectArr = Split(pickArr(sortedPos(i)), "//")
        For k = 1 To exampleNum(i)
            eg = Int((UBound(selectArr) - 1) * Rnd + 1)
            timesRec = 1
            Do While in_arr(selectArr(eg), "pp") > 0 Or Len(selectArr(eg)) >= 15
                If timesRec > 10 Then Exit Do
                eg = Int((UBound(selectArr) - 1) * Rnd + 1)
                timesRec = timesRec + 1
            Loop
            tblArr(row, col) = tblArr(row, col) + " //" + selectArr(eg)
        Next
    Next
End Function

Function cut_cake(ByVal cake As Integer, ByRef countArr)
    Dim i As Integer, total As Integer, maxPos As Integer, pieceSum As Integer
    For i = LBound(countArr) To UBound(countArr)
        total = total + countArr(i)
    Next
    Dim piece() As Integer
    ReDim piece(LBound(countArr) To UBound(countArr))
    For i = LBound(countArr) To UBound(countArr)
        piece(i) = CInt(cake * countArr(i) / total + 0.5)
        If piece(i) = 0 Then
            maxPos = i - 1
            Exit For
        End If
    Next
    Do While maxPos > cake
        piece(maxPos) = 0
        maxPos = maxPos - 1
    Loop
    For i = LBound(countArr) To maxPos
        pieceSum = pieceSum + piece(i)
    Next
    If pieceSum > cake And piece(LBound(countArr)) > int(cake/2) Then
        piece(LBound(countArr)) = piece(LBound(countArr)) - 1
        pieceSum = pieceSum - 1
    End If
    Do While pieceSum > cake
        If piece(maxPos) > 0 Then
            piece(maxPos) = piece(maxPos) - 1
            pieceSum = pieceSum - 1
            Else
                maxPos = maxPos - 1
                piece(maxPos) = piece(maxPos) - 1
                pieceSum = pieceSum - 1
        End If
    Loop
    cut_cake = piece
End Function

Public Function tone_example_table(ByVal cake As Integer, ByVal csToneCount, ByRef tblArr, ByRef initType, ByRef fnlType, ByRef hpTbl)
    Dim x As Integer, y As Integer, z As Integer, i As Integer, k As Integer, contrastRec As Integer
    Dim cRecNum As Integer, sRecNum As Integer, cRec(), sRec(), cMaxContrastRec As Integer, sMaxContrastRec As Integer
    ReDim cRec(1 To cake)
    ReDim sRec(1 To cake)
    csTag = csToneCount
    ssTag = UBound(tblArr) - csToneCount
    Do While sRecNum < cake And cRecNum < cake
        For x = LBound(initType) To UBound(initType)
            For y = LBound(fnlType) To UBound(fnlType)
                contrastRec = 0
                For z = LBound(tblArr) To UBound(tblArr)
                    If hpTbl(x, y, z) <> "" Then contrastRec = contrastRec + 1
                Next
                If fnl_num(fnlType(y)) \ 100000 >= 4 And fnl_num(fnlType(y)) \ 100000 <= 7 Then
                    If contrastRec > cMaxContrastRec Then cMaxContrastRec = contrastRec
                    If contrastRec = csTag And cRecNum <= cake - 1 Then
                        cRecNum = cRecNum + 1
                        cRec(cRecNum) = Array(x, y)
                    End If
                    Else
                        If contrastRec > sMaxContrastRec Then sMaxContrastRec = contrastRec
                        If contrastRec = ssTag And sRecNum <= cake - 1 Then
                            sRecNum = sRecNum + 1
                            sRec(sRecNum) = Array(x, y)
                        End If
                End If
            If cRecNum = cake And sRecNum = cake Then GoTo line1
            Next
        Next
        If csToneCount > 0 And cRecNum = 0 Then csTag = cMaxContrastRec
        If UBound(tblArr) - csToneCount > 0 And sRecNum = 0 Then ssTag = sMaxContrastRec
    Loop
line1: k = 1
    Dim rec(), hpstr As String
    ReDim rec(1 To cake * 2)
    If sRecNum > 0 Then Call fill_arr(rec, sRec, 1, 1, 1)
    If cRecNum > 0 Then Call fill_arr(rec, cRec, sRecNum + 1, 1, 1)
    For k = 1 To UBound(rec)
        If IsArray(rec(k)) Then
            For i = LBound(tblArr) To UBound(tblArr)
                hpstr = hpTbl(rec(k)(0), rec(k)(1), i)
                If hpstr <> "" Then
                    selectArr = Split(hpstr, "//")
                    eg = Int((UBound(selectArr) - 1) * Rnd + 1)
                    timesRec = 1
                    Do While in_arr(selectArr(eg), "pp") > 0 Or Len(selectArr(eg)) >= 15
                        If timesRec > 10 Then Exit Do
                        eg = Int((UBound(selectArr) - 1) * Rnd + 1)
                        timesRec = timesRec + 1
                    Loop
                    tblArr(i) = tblArr(i) + " //" + selectArr(eg)
                End If
            Next
        End If
    Next
    
End Function