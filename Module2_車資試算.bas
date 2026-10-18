Attribute VB_Name = "Module2"
Option Explicit

'==================================================================================
'  車資試算 － 楊梅火車站 → 目的地（火車）
'
'  核銷原則（票價以「費率對照」【三】現值為準；「交通費明細.pdf」僅供參考）：
'    ‧起點：楊梅火車站
'    ‧北上：楊梅→中壢採區間車（24元），中壢以北各站採自強號
'    ‧南下：楊梅→新竹採區間車（64元），新竹以南各站採自強號
'    ‧未達中壢／新竹之短程：一律以區間車核銷
'
'  用法：
'    1. 執行【建立車資試算表】→ 產生「車資試算」工作表（公式即時試算）
'    2. B5 選賽事場館（會自動帶出下車站與接駁費），或直接在 B6 填目的地火車站
'    3. 按【寫入費率對照【三】】→ 寫回「費率對照」【三】路線車資表
'
'  資料來源（優先序）：
'    1. 「費率對照」【三】現有路線明細 －－ 臺北89／臺中278／新營619、區間車 24／64／18
'    2. 使用者提供 －－ 中壢→基隆 222
'    3. 「交通費明細.pdf」 －－ 僅用於【三】沒有的目的地（樹林／板橋／松山／苗栗／
'       嘉義／臺南／新左營），該 pdf 的票價基準與【三】不同，來源欄已標示，請查證。
'    程式不代為推估任何票價；查不到票價時一律以 0 計並在 B25 顯示「※」。
'==================================================================================

Private Const SHEET_CALC As String = "車資試算"
Private Const SHEET_RATE As String = "費率對照"
Private Const MK_FARE    As String = "【三】"
Private Const FALLBACK_FARE_KEY As String = "＊預設"

' 輸入區
Private Const R_VENUE   As Long = 5     ' 賽事場館（選填，查【參考D】）
Private Const R_DEST    As Long = 6     ' 目的地火車站（留空＝依場館帶出）
Private Const R_TYPE    As Long = 7     ' 行程型態（留空＝自動判斷）
Private Const R_LOCAL   As Long = 8     ' 區間車票價（區間車直達用）
Private Const R_EXP     As Long = 9     ' 自強號票價
Private Const R_SHUT    As Long = 10    ' 在地接駁費
Private Const R_SHUTTXT As Long = 11    ' 在地接駁說明
Private Const R_KEY     As Long = 12    ' 寫入【三】的地點關鍵字
Private Const R_MEMO    As Long = 13    ' 地點補充說明
Private Const R_VAR     As Long = 14    ' 同時新增「臺／台」異體字列

' 試算結果區
Private Const R_STATION As Long = 17    ' 採用下車站
Private Const R_TYPEOUT As Long = 18
Private Const R_BOARD   As Long = 19
Private Const R_L1      As Long = 20
Private Const R_L2      As Long = 21
Private Const R_L3      As Long = 22
Private Const R_TOTAL   As Long = 23
Private Const R_DETAIL  As Long = 24
Private Const R_CHECK   As Long = 25
Private Const R_BTN     As Long = 27

' 參數區
Private Const R_P_FROM  As Long = 30
Private Const R_P_NORTH As Long = 31
Private Const R_P_SOUTH As Long = 32
Private Const R_P_F1N   As Long = 33
Private Const R_P_F1S   As Long = 34

' 參考表（A:C 參考A／E:H 參考B／J:L 參考C／N:S 參考D）
Private Const R_REFTITLE As Long = 36
Private Const R_REFHEAD  As Long = 37
Private Const R_REF      As Long = 38
Private Const R_REF_END  As Long = 200

Private Const T_LOCAL As String = "區間車直達"
Private Const T_NORTH As String = "北上"
Private Const T_SOUTH As String = "南下"


'==================================================================================
'  進入點
'==================================================================================
Public Sub 建立車資試算表()
    BuildCalcSheet False
End Sub


' 供自動化（PowerShell／COM）呼叫：不跳任何對話框，直接重建
Public Sub 建立車資試算表_靜默()
    BuildCalcSheet True
End Sub


Private Sub BuildCalcSheet(silent As Boolean)
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    If SheetExists(SHEET_CALC) Then
        If Not silent Then
            If MsgBox("「" & SHEET_CALC & "」已存在。" & vbLf & vbLf & _
                      "重建會覆蓋您在該表上的所有修改，" & vbLf & _
                      "包含【參考B】【參考C】【參考D】自行新增的票價。" & vbLf & vbLf & _
                      "確定要重建嗎？", vbYesNo + vbExclamation, "重建車資試算表") <> vbYes Then
                ThisWorkbook.Worksheets(SHEET_CALC).Activate
                Exit Sub
            End If
        End If
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SHEET_CALC).Delete
        Application.DisplayAlerts = True
    End If

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_CALC

    BuildLayout ws
    BuildFormulas ws
    BuildRefTables ws
    BuildFormat ws
    BuildButtons ws
    FillSample ws

    Application.ScreenUpdating = True

    If Not silent Then
        ws.Range("B" & R_VENUE).Select
        MsgBox "已建立「" & SHEET_CALC & "」工作表。" & vbLf & vbLf & _
               "‧B5 下拉選賽事場館，會自動帶出下車站與接駁費" & vbLf & _
               "‧或留空 B5，直接在 B6 填目的地火車站" & vbLf & _
               "‧票價查不到時請至臺鐵官網查證後填入 B8／B9" & vbLf & _
               "‧確認無誤後按【寫入費率對照【三】】", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not silent Then
        MsgBox "發生錯誤：" & vbLf & Err.Number & " － " & Err.Description, vbCritical, "錯誤"
    Else
        Err.Raise Err.Number, "BuildCalcSheet", Err.Description
    End If
End Sub


' 清空輸入區（保留參數與參考表）
Public Sub 清除車資試算輸入()
    Dim ws As Worksheet
    If Not SheetExists(SHEET_CALC) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_CALC)
    ws.Range("B" & R_VENUE & ":B" & R_MEMO).ClearContents
    ws.Range("B" & R_VENUE).Select
End Sub


'==================================================================================
'  寫回「費率對照」【三】
'==================================================================================
Public Sub 寫入費率對照()
    WriteBack False
End Sub


' 供自動化（PowerShell／COM）呼叫：不跳任何對話框。
' 檢查未通過（B25 出現「※」）時一律不寫入；覆寫既有列時直接覆寫。
Public Sub 寫入費率對照_靜默()
    WriteBack True
End Sub


Private Sub WriteBack(silent As Boolean)
    Dim wsC As Worksheet, wsR As Worksheet
    Dim key As String, altKey As String, detail As String, memo As String
    Dim venue As String, station As String, chk As String
    Dim fare As Double
    Dim act1 As String, act2 As String, msg As String

    On Error GoTo ErrHandler

    If Not SheetExists(SHEET_CALC) Then
        If Not silent Then MsgBox "找不到「" & SHEET_CALC & "」工作表，請先執行【建立車資試算表】。", _
                                  vbExclamation, "無法寫入"
        Exit Sub
    End If
    If Not SheetExists(SHEET_RATE) Then
        If Not silent Then MsgBox "找不到「" & SHEET_RATE & "」工作表。", vbExclamation, "無法寫入"
        Exit Sub
    End If

    Set wsC = ThisWorkbook.Worksheets(SHEET_CALC)
    Set wsR = ThisWorkbook.Worksheets(SHEET_RATE)

    venue = Trim$(CStr(wsC.Range("B" & R_VENUE).Value))
    station = Trim$(CStr(wsC.Range("B" & R_STATION).Value))
    key = Trim$(CStr(wsC.Range("B" & R_KEY).Value))
    If key = "" Then key = station
    detail = CStr(wsC.Range("B" & R_DETAIL).Value)
    memo = Trim$(CStr(wsC.Range("B" & R_MEMO).Value))
    If memo = "" And venue <> "" Then memo = LookupRefD(wsC, venue, 5)
    fare = NumOf(wsC.Range("B" & R_TOTAL).Value)
    chk = CStr(wsC.Range("B" & R_CHECK).Value)

    If key = "" Then
        If Not silent Then MsgBox "請先於 B" & R_VENUE & " 選賽事場館，或於 B" & R_DEST & " 填目的地火車站。", _
                                  vbExclamation, "無法寫入"
        Exit Sub
    End If
    If key = FALLBACK_FARE_KEY Then
        If Not silent Then MsgBox "地點關鍵字不可為「" & FALLBACK_FARE_KEY & "」。", vbExclamation, "無法寫入"
        Exit Sub
    End If
    If fare <= 0 Then
        If Not silent Then MsgBox "單趟票價合計為 0，無法寫入。" & vbLf & vbLf & chk, vbExclamation, "無法寫入"
        Exit Sub
    End If
    If InStr(1, chk, "※") > 0 Then
        ' 靜默模式一律不寫入，避免把未查證的票價寫進【三】
        If silent Then Exit Sub
        If MsgBox(chk & vbLf & vbLf & "仍要寫入嗎？", vbYesNo + vbExclamation, "檢查未通過") <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False

    act1 = WriteFareRow(wsR, key, fare, detail, memo, silent)

    ' 「臺／台」異體字：一併新增，與原表既有作法一致
    If InStr(1, CStr(wsC.Range("B" & R_VAR).Value), "是", vbTextCompare) > 0 Then
        altKey = VariantKey(key)
        If altKey <> "" Then act2 = WriteFareRow(wsR, altKey, fare, detail, memo, silent)
    End If

    Application.ScreenUpdating = True

    If act1 = "" Then Exit Sub
    If silent Then Exit Sub

    msg = "已" & act1 & "「費率對照」【三】：" & key & "　" & Format$(fare, "#,##0") & " 元"
    If act2 <> "" Then msg = msg & vbLf & "已" & act2 & "異體字列：" & altKey
    msg = msg & vbLf & vbLf & "請確認【三】的路線明細與票價圖是否齊備。"

    wsR.Activate
    MsgBox msg, vbInformation, "寫入完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    If Not silent Then
        MsgBox "發生錯誤：" & vbLf & Err.Number & " － " & Err.Description, vbCritical, "錯誤"
    Else
        Err.Raise Err.Number, "WriteBack", Err.Description
    End If
End Sub


' 從【參考D】取某場館的第 col 欄（1=場館 2=下車站 3=接駁費 4=接駁說明 5=地點補充）
Private Function LookupRefD(wsC As Worksheet, venue As String, col As Long) As String
    Dim r As Long
    LookupRefD = ""
    For r = R_REF To R_REF_END
        If Trim$(CStr(wsC.Cells(r, 14).Value)) = "" Then Exit Function
        If StrComp(Trim$(CStr(wsC.Cells(r, 14).Value)), venue, vbTextCompare) = 0 Then
            LookupRefD = Trim$(CStr(wsC.Cells(r, 13 + col).Value))
            Exit Function
        End If
    Next r
End Function


' 寫入或更新【三】的一列，回傳「新增」或「更新」（取消時回傳空字串）
Private Function WriteFareRow(wsR As Worksheet, key As String, fare As Double, _
                              detail As String, memo As String, silent As Boolean) As String
    Dim r As Long, insertAt As Long, hitRow As Long
    Dim oldFare As Double

    WriteFareRow = ""
    r = BlockFirstRow(wsR, MK_FARE)
    If r = 0 Then
        If Not silent Then MsgBox "「" & SHEET_RATE & "」中找不到區塊標記「" & MK_FARE & "」。", _
                                  vbCritical, "無法寫入"
        Exit Function
    End If

    hitRow = 0
    insertAt = 0
    Do While Trim$(CStr(wsR.Cells(r, 1).Value)) <> ""
        If Trim$(CStr(wsR.Cells(r, 1).Value)) = FALLBACK_FARE_KEY Then
            insertAt = r
            Exit Do
        End If
        If StrComp(Trim$(CStr(wsR.Cells(r, 1).Value)), key, vbTextCompare) = 0 Then
            hitRow = r
            Exit Do
        End If
        r = r + 1
    Loop
    If insertAt = 0 And hitRow = 0 Then insertAt = r   ' 沒有＊預設列時，接在區塊最後

    If hitRow > 0 Then
        oldFare = NumOf(wsR.Cells(hitRow, 2).Value)
        If oldFare > 0 And oldFare <> fare And Not silent Then
            If MsgBox("【三】已有「" & key & "」，單趟票價 " & Format$(oldFare, "#,##0") & " 元。" & vbLf & vbLf & _
                      "要覆寫為 " & Format$(fare, "#,##0") & " 元嗎？", _
                      vbYesNo + vbQuestion, "覆寫確認") <> vbYes Then Exit Function
        End If
        r = hitRow
        WriteFareRow = "更新"
    Else
        wsR.Rows(insertAt).Insert Shift:=xlDown
        r = insertAt
        WriteFareRow = "新增"
    End If

    wsR.Cells(r, 1).Value = key
    wsR.Cells(r, 2).Value = fare
    wsR.Cells(r, 3).Value = detail
    wsR.Cells(r, 4).Value = memo

    With wsR.Range(wsR.Cells(r, 1), wsR.Cells(r, 4))
        .Interior.Pattern = xlNone          ' 清掉「待填」黃底
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    wsR.Cells(r, 2).NumberFormat = "#,##0"
    wsR.Rows(r).EntireRow.AutoFit
End Function


' 臺 ←→ 台 互換；無對應時回傳空字串
Private Function VariantKey(key As String) As String
    Dim s As String
    VariantKey = ""
    If InStr(1, key, "臺") > 0 Then
        s = Replace(key, "臺", "台")
    ElseIf InStr(1, key, "台") > 0 Then
        s = Replace(key, "台", "臺")
    Else
        Exit Function
    End If
    If s <> key Then VariantKey = s
End Function


'==================================================================================
'  建表
'==================================================================================
Private Sub BuildLayout(ws As Worksheet)
    With ws
        .Range("A1").Value = "【車資試算】楊梅火車站 → 目的地　（票價以「費率對照」【三】現值為準；結果可一鍵寫入【三】）"
        .Range("A2").Value = "※ 核銷原則：北上 楊梅→中壢採區間車24元、中壢以北採自強號；" & _
                             "南下 楊梅→新竹採區間車64元、新竹以南採自強號；未達中壢／新竹之短程一律區間車。"

        .Range("A4").Value = "■ 輸入"
        .Cells(R_VENUE, 1).Value = "賽事場館（選填，下拉選單）"
        .Cells(R_VENUE, 3).Value = "選了會自動帶出下車站、接駁費與接駁說明（見【參考D】）"
        .Cells(R_DEST, 1).Value = "目的地火車站"
        .Cells(R_DEST, 3).Value = "留空＝依 B" & R_VENUE & " 場館帶出；也可直接填站名（例：臺北／苗栗／新營）"
        .Cells(R_TYPE, 1).Value = "行程型態（留空＝自動判斷）"
        .Cells(R_TYPE, 3).Value = "區間車直達／北上／南下；自動判斷結果見 B" & R_TYPEOUT
        .Cells(R_LOCAL, 1).Value = "區間車票價（元）"
        .Cells(R_LOCAL, 3).Value = "僅「區間車直達」用；留空＝查【參考C】"
        .Cells(R_EXP, 1).Value = "自強號票價（元）"
        .Cells(R_EXP, 3).Value = "僅北上／南下用；留空＝查【參考B】。查無請至臺鐵官網查證後填入"
        .Cells(R_SHUT, 1).Value = "在地接駁費（元／單趟）"
        .Cells(R_SHUT, 3).Value = "留空＝依場館查【參考D】；沒有接駁請填 0"
        .Cells(R_SHUTTXT, 1).Value = "在地接駁說明"
        .Cells(R_SHUTTXT, 3).Value = "留空＝依場館查【參考D】，再查不到則顯示「(公車)」"
        .Cells(R_KEY, 1).Value = "寫入【三】的地點關鍵字"
        .Cells(R_KEY, 3).Value = "留空＝同 B" & R_STATION & " 下車站。此關鍵字會與「競賽資訊」比賽地點做包含比對"
        .Cells(R_MEMO, 1).Value = "地點補充說明（寫入【三】D欄）"
        .Cells(R_MEMO, 3).Value = "留空＝依場館查【參考D】"
        .Cells(R_VAR, 1).Value = "同時新增「臺／台」異體字列"
        .Cells(R_VAR, 2).Value = "是"
        .Cells(R_VAR, 3).Value = "例：寫入「臺南」時一併寫入「台南」，與原表既有作法一致"

        .Range("A16").Value = "■ 試算結果（公式自動計算）"
        .Cells(R_STATION, 1).Value = "採用下車站"
        .Cells(R_TYPEOUT, 1).Value = "採用行程型態"
        .Cells(R_BOARD, 1).Value = "自強號上車站"
        .Cells(R_L1, 1).Value = "第1段　區間車（元）"
        .Cells(R_L2, 1).Value = "第2段　自強號（元）"
        .Cells(R_L3, 1).Value = "第3段　在地接駁（元）"
        .Cells(R_TOTAL, 1).Value = "單趟票價合計（元）"
        .Cells(R_DETAIL, 1).Value = "路線明細（寫入【三】C欄）"
        .Cells(R_CHECK, 1).Value = "檢查"

        .Range("A29").Value = "■ 參數（依「費率對照」【三】現有路線明細；可自行修改）"
        .Cells(R_P_FROM, 1).Value = "出發站"
        .Cells(R_P_FROM, 2).Value = "楊梅"
        .Cells(R_P_NORTH, 1).Value = "北上　自強號上車站"
        .Cells(R_P_NORTH, 2).Value = "中壢"
        .Cells(R_P_SOUTH, 1).Value = "南下　自強號上車站"
        .Cells(R_P_SOUTH, 2).Value = "新竹"
        .Cells(R_P_F1N, 1).Value = "區間車　楊梅→中壢（元）"
        .Cells(R_P_F1N, 2).Value = 24
        .Cells(R_P_F1N, 3).Value = "【三】臺北路線明細之第1段（pdf 為 18，未採用）"
        .Cells(R_P_F1S, 1).Value = "區間車　楊梅→新竹（元）"
        .Cells(R_P_F1S, 2).Value = 64
        .Cells(R_P_F1S, 3).Value = "【三】新營／臺中路線明細之第1段（pdf 為 37，未採用）"
    End With
End Sub


Private Sub BuildFormulas(ws As Worksheet)
    Dim refA As String, refB As String, refC As String, refD As String
    refA = "$A$" & R_REF & ":$C$" & R_REF_END
    refB = "$E$" & R_REF & ":$G$" & R_REF_END
    refC = "$J$" & R_REF & ":$K$" & R_REF_END
    refD = "$N$" & R_REF & ":$S$" & R_REF_END

    With ws
        .Cells(R_STATION, 2).Formula = _
            "=IF($B$" & R_DEST & "<>"""",TRIM($B$" & R_DEST & ")," & _
            "IFERROR(VLOOKUP($B$" & R_VENUE & "," & refD & ",2,FALSE),""""))"

        .Cells(R_TYPEOUT, 2).Formula = _
            "=IF($B$" & R_TYPE & "<>"""",TRIM($B$" & R_TYPE & ")," & _
            "IFERROR(VLOOKUP($B$" & R_STATION & "," & refA & ",2,FALSE),""？無法判斷，請於 B" & R_TYPE & " 指定""))"

        .Cells(R_BOARD, 2).Formula = _
            "=IF($B$" & R_TYPEOUT & "=""" & T_NORTH & """,$B$" & R_P_NORTH & _
            ",IF($B$" & R_TYPEOUT & "=""" & T_SOUTH & """,$B$" & R_P_SOUTH & ",""（不搭自強號）""))"

        .Cells(R_L1, 2).Formula = _
            "=IF($B$" & R_TYPEOUT & "=""" & T_NORTH & """,$B$" & R_P_F1N & _
            ",IF($B$" & R_TYPEOUT & "=""" & T_SOUTH & """,$B$" & R_P_F1S & _
            ",IF($B$" & R_LOCAL & "<>"""",$B$" & R_LOCAL & _
            ",IFERROR(VLOOKUP($B$" & R_STATION & "," & refC & ",2,FALSE),0))))"

        .Cells(R_L1, 3).Formula = _
            "=IF(N($B$" & R_L1 & ")<=0,"""",""(區間)""&$B$" & R_P_FROM & "&""-""&" & _
            "IF($B$" & R_TYPEOUT & "=""" & T_NORTH & """,$B$" & R_P_NORTH & _
            ",IF($B$" & R_TYPEOUT & "=""" & T_SOUTH & """,$B$" & R_P_SOUTH & ",$B$" & R_STATION & "))" & _
            "&TEXT($B$" & R_L1 & ",""0"")&""元"")"

        .Cells(R_L2, 2).Formula = _
            "=IF(OR($B$" & R_TYPEOUT & "=""" & T_NORTH & """,$B$" & R_TYPEOUT & "=""" & T_SOUTH & """)," & _
            "IF($B$" & R_EXP & "<>"""",$B$" & R_EXP & _
            ",IFERROR(VLOOKUP($B$" & R_STATION & "," & refB & ",3,FALSE),0)),0)"

        .Cells(R_L2, 3).Formula = _
            "=IF(N($B$" & R_L2 & ")<=0,"""",""(自強)""&$B$" & R_BOARD & "&""-""&$B$" & R_STATION & _
            "&TEXT($B$" & R_L2 & ",""0"")&""元"")"

        .Cells(R_L3, 2).Formula = _
            "=IF($B$" & R_SHUT & "<>"""",$B$" & R_SHUT & _
            ",IFERROR(VLOOKUP($B$" & R_VENUE & "," & refD & ",3,FALSE),0))"

        .Cells(R_L3, 3).Formula = _
            "=IF(N($B$" & R_L3 & ")<=0,""""," & _
            "IF($B$" & R_SHUTTXT & "<>"""",$B$" & R_SHUTTXT & _
            ",IFERROR(VLOOKUP($B$" & R_VENUE & "," & refD & ",4,FALSE),""(公車)""))" & _
            "&TEXT($B$" & R_L3 & ",""0"")&""元"")"

        .Cells(R_TOTAL, 2).Formula = "=$B$" & R_L1 & "+$B$" & R_L2 & "+$B$" & R_L3

        .Cells(R_DETAIL, 2).Formula = _
            "=MID(IF($C$" & R_L1 & "="""","""",CHAR(10)&$C$" & R_L1 & ")" & _
            "&IF($C$" & R_L2 & "="""","""",CHAR(10)&$C$" & R_L2 & ")" & _
            "&IF($C$" & R_L3 & "="""","""",CHAR(10)&$C$" & R_L3 & "),2,999)"

        .Cells(R_CHECK, 2).Formula = _
            "=IF($B$" & R_STATION & "=""""," & _
            """※ 請於 B" & R_VENUE & " 選場館，或於 B" & R_DEST & " 填目的地火車站。""," & _
            "IF(LEFT($B$" & R_TYPEOUT & ",1)=""？"",""※ 【參考A】查無「""&$B$" & R_STATION & _
            "&""」，請於 B" & R_TYPE & " 指定行程型態，或到【參考A】新增該站。""," & _
            "IF(AND(OR($B$" & R_TYPEOUT & "=""" & T_NORTH & """,$B$" & R_TYPEOUT & "=""" & T_SOUTH & """)," & _
            "$B$" & R_L2 & "<=0),""※ 查無自強號票價，請至臺鐵官網查證後填入 B" & R_EXP & "（或新增至【參考B】）。""," & _
            "IF(AND($B$" & R_TYPEOUT & "=""" & T_LOCAL & """,$B$" & R_L1 & "<=0)," & _
            """※ 查無區間車票價，請填入 B" & R_LOCAL & "（或新增至【參考C】）。""," & _
            "IF($B$" & R_TOTAL & "<=0,""※ 合計為 0，無法寫入。""," & _
            """○ 試算完成，可按【寫入費率對照【三】】。"")))))"
    End With
End Sub


Private Sub BuildRefTables(ws As Worksheet)
    Dim rows_ As Variant, p As Variant
    Dim i As Long, r As Long, c As Long

    With ws
        '--- 【參考A】站名 → 行程型態 -------------------------------------------
        .Cells(R_REFTITLE, 1).Value = "■【參考A】站名 → 行程型態　（完全比對；可自行增修）"
        .Cells(R_REFHEAD, 1).Value = "站名"
        .Cells(R_REFHEAD, 2).Value = "行程型態"
        .Cells(R_REFHEAD, 3).Value = "備註"
        FillBlock ws, RefAData(), 1, 3

        '--- 【參考B】自強號票價 -------------------------------------------------
        .Cells(R_REFTITLE, 5).Value = "■【參考B】自強號票價　（優先以【三】明細為準；【三】未涵蓋者暫取 pdf，見來源欄）"
        .Cells(R_REFHEAD, 5).Value = "目的地站"
        .Cells(R_REFHEAD, 6).Value = "上車站"
        .Cells(R_REFHEAD, 7).Value = "票價(元)"
        .Cells(R_REFHEAD, 8).Value = "來源／備註"
        FillBlock ws, RefBData(), 5, 4

        '--- 【參考C】區間車票價 -------------------------------------------------
        .Cells(R_REFTITLE, 10).Value = "■【參考C】區間車票價（楊梅出發）　（以【三】現值為準）"
        .Cells(R_REFHEAD, 10).Value = "目的地站"
        .Cells(R_REFHEAD, 11).Value = "票價(元)"
        .Cells(R_REFHEAD, 12).Value = "來源／備註"
        FillBlock ws, RefCData(), 10, 3

        '--- 【參考D】賽事場館對照 -----------------------------------------------
        .Cells(R_REFTITLE, 14).Value = "■【參考D】賽事場館 → 下車站／在地接駁　（交通費明細.pdf ＋ 運動賽事地點.docx）"
        .Cells(R_REFHEAD, 14).Value = "賽事場館"
        .Cells(R_REFHEAD, 15).Value = "下車站"
        .Cells(R_REFHEAD, 16).Value = "接駁費(元)"
        .Cells(R_REFHEAD, 17).Value = "接駁說明"
        .Cells(R_REFHEAD, 18).Value = "地點補充說明"
        .Cells(R_REFHEAD, 19).Value = "來源／備註"
        FillBlock ws, RefDData(), 14, 6

        ' 數字欄轉成數值
        For r = R_REF To R_REF_END
            If Trim$(CStr(.Cells(r, 7).Value)) <> "" Then .Cells(r, 7).Value = NumOf(.Cells(r, 7).Value)
            If Trim$(CStr(.Cells(r, 11).Value)) <> "" Then .Cells(r, 11).Value = NumOf(.Cells(r, 11).Value)
            If Trim$(CStr(.Cells(r, 16).Value)) <> "" Then .Cells(r, 16).Value = NumOf(.Cells(r, 16).Value)
        Next r
    End With
End Sub


' 以 ; 分列、| 分欄，寫入 firstCol 起的 nCols 欄
Private Sub FillBlock(ws As Worksheet, data As String, firstCol As Long, nCols As Long)
    Dim rows_ As Variant, p As Variant
    Dim i As Long, j As Long, r As Long
    rows_ = Split(data, ";")
    r = R_REF
    For i = LBound(rows_) To UBound(rows_)
        If Trim$(CStr(rows_(i))) <> "" Then
            p = Split(CStr(rows_(i)), "|")
            For j = 0 To nCols - 1
                If j <= UBound(p) Then ws.Cells(r, firstCol + j).Value = p(j)
            Next j
            r = r + 1
        End If
    Next i
End Sub


' 站名｜行程型態｜備註
Private Function RefAData() As String
    Dim s As String
    s = "桃園|" & T_LOCAL & "|;內壢|" & T_LOCAL & "|;中壢|" & T_LOCAL & "|;埔心|" & T_LOCAL & "|;"
    s = s & "楊梅|" & T_LOCAL & "|;富岡|" & T_LOCAL & "|;湖口|" & T_LOCAL & "|;新豐|" & T_LOCAL & "|;"
    s = s & "竹北|" & T_LOCAL & "|;北新竹|" & T_LOCAL & "|;新竹|" & T_LOCAL & "|;鶯歌|" & T_LOCAL & "|;"
    s = s & "平鎮|" & T_LOCAL & "|無臺鐵站，就近中壢站;八德|" & T_LOCAL & "|無臺鐵站，就近中壢站;"
    s = s & "龜山|" & T_LOCAL & "|無臺鐵站，就近桃園站;蘆竹|" & T_LOCAL & "|無臺鐵站，就近桃園站;"
    s = s & "大園|" & T_LOCAL & "|無臺鐵站，就近桃園站;大溪|" & T_LOCAL & "|無臺鐵站，就近桃園站;"
    s = s & "龍潭|" & T_LOCAL & "|無臺鐵站，就近中壢站;新屋|" & T_LOCAL & "|無臺鐵站，就近楊梅站;"
    s = s & "觀音|" & T_LOCAL & "|無臺鐵站，就近楊梅站;"

    s = s & "山佳|" & T_NORTH & "|;樹林|" & T_NORTH & "|;板橋|" & T_NORTH & "|;萬華|" & T_NORTH & "|;"
    s = s & "臺北|" & T_NORTH & "|;台北|" & T_NORTH & "|;松山|" & T_NORTH & "|;南港|" & T_NORTH & "|;"
    s = s & "汐止|" & T_NORTH & "|;七堵|" & T_NORTH & "|;八堵|" & T_NORTH & "|;基隆|" & T_NORTH & "|;"
    s = s & "三重|" & T_NORTH & "|;新莊|" & T_NORTH & "|;瑞芳|" & T_NORTH & "|;福隆|" & T_NORTH & "|;"
    s = s & "淡水|" & T_NORTH & "|無臺鐵站，臺北轉捷運;"
    s = s & "宜蘭|" & T_NORTH & "|東部幹線經八堵;羅東|" & T_NORTH & "|東部幹線經八堵;蘇澳|" & T_NORTH & "|東部幹線經八堵;"
    s = s & "花蓮|" & T_NORTH & "|北迴線;玉里|" & T_NORTH & "|北迴線;池上|" & T_NORTH & "|北迴線;"
    s = s & "臺東|" & T_NORTH & "|北迴線；改走南迴請於 B7 指定南下;台東|" & T_NORTH & "|北迴線；改走南迴請於 B7 指定南下;"

    s = s & "香山|" & T_SOUTH & "|;竹南|" & T_SOUTH & "|;頭份|" & T_SOUTH & "|無臺鐵站，就近竹南站;"
    s = s & "造橋|" & T_SOUTH & "|;苗栗|" & T_SOUTH & "|;銅鑼|" & T_SOUTH & "|;三義|" & T_SOUTH & "|;"
    s = s & "后里|" & T_SOUTH & "|;豐原|" & T_SOUTH & "|;潭子|" & T_SOUTH & "|;臺中|" & T_SOUTH & "|;台中|" & T_SOUTH & "|;"
    s = s & "新烏日|" & T_SOUTH & "|;烏日|" & T_SOUTH & "|;大甲|" & T_SOUTH & "|海線;清水|" & T_SOUTH & "|海線;沙鹿|" & T_SOUTH & "|海線;"
    s = s & "彰化|" & T_SOUTH & "|;員林|" & T_SOUTH & "|;田中|" & T_SOUTH & "|;二水|" & T_SOUTH & "|;"
    s = s & "南投|" & T_SOUTH & "|縣內無臺鐵站，就近臺中／彰化站;斗六|" & T_SOUTH & "|;斗南|" & T_SOUTH & "|;"
    s = s & "虎尾|" & T_SOUTH & "|無臺鐵站，就近斗南站;民雄|" & T_SOUTH & "|;嘉義|" & T_SOUTH & "|;水上|" & T_SOUTH & "|;"
    s = s & "朴子|" & T_SOUTH & "|無臺鐵站，就近嘉義站;新營|" & T_SOUTH & "|;柳營|" & T_SOUTH & "|;善化|" & T_SOUTH & "|;"
    s = s & "新市|" & T_SOUTH & "|;永康|" & T_SOUTH & "|;臺南|" & T_SOUTH & "|;台南|" & T_SOUTH & "|;"
    s = s & "岡山|" & T_SOUTH & "|;橋頭|" & T_SOUTH & "|;楠梓|" & T_SOUTH & "|;"
    s = s & "新左營|" & T_SOUTH & "|;左營|" & T_SOUTH & "|;高雄|" & T_SOUTH & "|;鳳山|" & T_SOUTH & "|;"
    s = s & "屏東|" & T_SOUTH & "|;潮州|" & T_SOUTH & "|;枋寮|" & T_SOUTH & "|;"
    s = s & "恆春|" & T_SOUTH & "|無臺鐵站，枋寮轉乘;墾丁|" & T_SOUTH & "|無臺鐵站，枋寮轉乘"
    RefAData = s
End Function


' 目的地站｜上車站｜票價｜來源
Private Function RefBData() As String
    Dim s As String, Q As String, X As String
    Q = "【三】路線明細"
    X = "交通費明細.pdf（基準與【三】不同，請查證）"

    ' 一、可由【三】現有明細直接推得者
    s = "臺北|中壢|89|" & Q & ";台北|中壢|89|" & Q & "（異體字）;"
    s = s & "臺中|新竹|278|" & Q & ";台中|新竹|278|" & Q & "（異體字）;"
    s = s & "新營|新竹|619|" & Q & ";"
    ' 二、使用者提供（【三】基隆列 133 係借用臺北票價，非實際基隆票價）
    s = s & "基隆|中壢|222|使用者提供;"
    ' 三、【三】未涵蓋之目的地，暫取 pdf 值
    s = s & "樹林|中壢|64|" & X & ";板橋|中壢|78|" & X & ";松山|中壢|104|" & X & ";"
    s = s & "苗栗|新竹|69|" & X & ";嘉義|新竹|420|" & X & ";"
    s = s & "臺南|新竹|553|" & X & ";台南|新竹|553|" & X & ";新左營|新竹|609|" & X
    RefBData = s
End Function


' 目的地站｜票價｜來源
Private Function RefCData() As String
    Dim s As String, Q As String
    Q = "【三】現值"
    s = "中壢|18|" & Q & "（pdf 亦為 18）;竹北|64|" & Q & "（pdf 為 34，未採用）;"
    s = s & "新竹|64|" & Q & "（pdf 為 37，未採用）"
    RefCData = s
End Function


' 賽事場館｜下車站｜接駁費｜接駁說明｜地點補充說明｜來源
Private Function RefDData() As String
    Dim s As String, P As String, D As String
    P = "交通費明細.pdf"
    D = "運動賽事地點.docx（票價待查證）"

    s = "新竹縣體育館|竹北|15|(市區公車)|新竹縣竹北市；市民公車0元、市區公車15元|" & P & ";"
    s = s & "板橋第一體育場|板橋|15|(市區公車)|新北市板橋區；步行約10分鐘可0元|" & P & ";"
    s = s & "板橋第一運動場|板橋|15|(市區公車)|新北市板橋區；步行約10分鐘可0元|" & P & ";"
    s = s & "板橋體育館|板橋|15|(市區公車)|新北市板橋區中正路；步行約15分鐘可0元|" & P & ";"
    s = s & "板樹體育館|樹林|15|(公車藍38)樹林後站-板樹體育館|新北市溪城路；步行約15分鐘可0元|" & P & ";"
    s = s & "苗栗巨蛋體育館|苗栗|25|(公路客運)|苗栗市；步行約15分鐘可0元|" & P & ";"
    s = s & "高雄國家體育場|新左營|20|(高雄捷運)新左營-世運站|世運主場館|" & P & ";"
    s = s & "世運主場館|新左營|20|(高雄捷運)新左營-世運站|高雄國家體育場|" & P & ";"
    s = s & "臺北田徑場|臺北|20|(臺北捷運)臺北-台北小巨蛋站|臺北小巨蛋旁；亦可搭至松山站(自強104元)|" & P & ";"
    s = s & "臺南市立體育場|臺南|0||臺南市南區；市區公車／計程車核實報支|" & P & " 接駁待查;"
    s = s & "新營體育館|新營|26|(大台南公車棕3線)新營車站-新營體育館|臺南市新營區長榮路二段78號(約217公里)|【三】現值（pdf 為黃幹線18元）;"
    s = s & "嘉義縣立體育館|嘉義|45|(BRT 7211)嘉義後站-體育館|嘉義縣朴子市|" & P & ";"

    s = s & "嘉義縣立田徑場|嘉義|45|(BRT 7211)嘉義後站-體育館|嘉義縣朴子市；115全中運田徑|" & P & "接駁＋" & D & ";"
    s = s & "朴子國中體育館|嘉義|45|(BRT 7211)嘉義後站-朴子|嘉義縣朴子市；115全中運角力|" & D & ";"
    s = s & "臺南市立田徑場|臺南|0||114全中運田徑|" & D & ";"
    s = s & "成功大學光復校區中正堂|臺南|0||114全中運柔道|" & D & ";"
    s = s & "臺中市立豐原體育場|豐原|0||114全中錦田徑|" & D & ";"
    s = s & "臺中港區體育館|沙鹿|0||跆拳道；海線，亦可自清水站|" & D & ";"
    s = s & "臺中市中港體育館|臺中|0||柔道|" & D & ";"
    s = s & "花蓮縣立體育館|花蓮|0||小巨蛋；跆拳道|" & D & ";"
    s = s & "臺北體育館|臺北|20|(臺北捷運)|紅館；角力／柔道|" & D & ";"
    s = s & "臺北市立體育館|臺北|20|(臺北捷運)|柔道|" & D & ";"
    s = s & "臺北市立大同高中體育館|臺北|20|(臺北捷運)|113全中運柔道|" & D & ";"
    s = s & "臺北市立大學天母校區體育館|臺北|20|(臺北捷運＋公車)|113全中運跆拳道|" & D & ";"
    s = s & "文化大學體育館|臺北|20|(臺北捷運＋公車)|柔道|" & D & ";"
    s = s & "嘉義市立體育館|嘉義|0||角力|" & D
    RefDData = s
End Function


Private Sub BuildFormat(ws As Worksheet)
    Dim lastVenue As Long
    With ws
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 52
        .Columns("D").ColumnWidth = 3
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 10
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 24
        .Columns("I").ColumnWidth = 3
        .Columns("J").ColumnWidth = 12
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 26
        .Columns("M").ColumnWidth = 3
        .Columns("N").ColumnWidth = 30
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 11
        .Columns("Q").ColumnWidth = 34
        .Columns("R").ColumnWidth = 40
        .Columns("S").ColumnWidth = 30

        .Range("A1:C1").Merge
        With .Range("A1")
            .Font.Size = 14
            .Font.Bold = True
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        .Rows(1).RowHeight = 32

        .Range("A2:C2").Merge
        With .Range("A2")
            .Font.Color = RGB(150, 0, 0)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(2).RowHeight = 34

        .Range("A4").Font.Bold = True
        .Range("A16").Font.Bold = True
        .Range("A29").Font.Bold = True
        .Cells(R_REFTITLE, 1).Font.Bold = True
        .Cells(R_REFTITLE, 5).Font.Bold = True
        .Cells(R_REFTITLE, 10).Font.Bold = True
        .Cells(R_REFTITLE, 14).Font.Bold = True

        ' 輸入區
        With .Range("B" & R_VENUE & ":B" & R_VAR)
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With
        .Range("A" & R_VENUE & ":A" & R_VAR).Font.Bold = True
        .Range("C" & R_VENUE & ":C" & R_VAR).Font.Color = RGB(110, 110, 110)
        .Range("B" & R_LOCAL & ":B" & R_SHUT).NumberFormat = "#,##0"

        ' 結果區
        With .Range("B" & R_STATION & ":B" & R_CHECK)
            .Interior.Color = RGB(217, 234, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With
        .Range("A" & R_STATION & ":A" & R_CHECK).Font.Bold = True
        .Range("B" & R_L1 & ":B" & R_TOTAL).NumberFormat = "#,##0"
        With .Range("B" & R_TOTAL)
            .Font.Bold = True
            .Font.Size = 12
        End With
        With .Range("B" & R_DETAIL)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        .Rows(R_DETAIL).RowHeight = 58
        .Range("B" & R_CHECK).WrapText = True
        .Rows(R_CHECK).RowHeight = 32
        .Range("C" & R_L1 & ":C" & R_L3).Font.Color = RGB(110, 110, 110)

        ' 參數區
        With .Range("B" & R_P_FROM & ":B" & R_P_F1S)
            .Interior.Color = RGB(235, 235, 235)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With

        ' 參考表表頭
        FormatHeader ws, .Range(.Cells(R_REFHEAD, 1), .Cells(R_REFHEAD, 3))
        FormatHeader ws, .Range(.Cells(R_REFHEAD, 5), .Cells(R_REFHEAD, 8))
        FormatHeader ws, .Range(.Cells(R_REFHEAD, 10), .Cells(R_REFHEAD, 12))
        FormatHeader ws, .Range(.Cells(R_REFHEAD, 14), .Cells(R_REFHEAD, 19))

        ' 場館下拉
        lastVenue = .Cells(R_REF_END, 14).End(xlUp).Row
        If lastVenue < R_REF Then lastVenue = R_REF
        With .Range("B" & R_VENUE).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Operator:=xlBetween, Formula1:="=$N$" & R_REF & ":$N$" & lastVenue
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        With .Range("B" & R_TYPE).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=T_LOCAL & "," & T_NORTH & "," & T_SOUTH
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        With .Range("B" & R_VAR).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="是,否"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With
End Sub


Private Sub FormatHeader(ws As Worksheet, rg As Range)
    With rg
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
End Sub


Private Sub BuildButtons(ws As Worksheet)
    Dim b As Object
    ws.Rows(R_BTN).RowHeight = 34

    Set b = ws.Buttons.Add(ws.Cells(R_BTN, 1).Left + 2, ws.Cells(R_BTN, 1).Top + 3, 200, 26)
    b.OnAction = "寫入費率對照"
    b.Caption = "寫入費率對照【三】"
    b.Font.Size = 11

    Set b = ws.Buttons.Add(ws.Cells(R_BTN, 2).Left + 2, ws.Cells(R_BTN, 2).Top + 3, 110, 26)
    b.OnAction = "清除車資試算輸入"
    b.Caption = "清除輸入"
    b.Font.Size = 11

    Set b = ws.Buttons.Add(ws.Cells(R_BTN, 3).Left + 2, ws.Cells(R_BTN, 3).Top + 3, 160, 26)
    b.OnAction = "計算差旅費"
    b.Caption = "重算所有概算表"
    b.Font.Size = 11
End Sub


' 預帶一組範例（新營體育館），讓使用者一眼看懂欄位怎麼填
Private Sub FillSample(ws As Worksheet)
    ws.Cells(R_VENUE, 2).Value = "新營體育館"
End Sub


'==================================================================================
'  工具
'==================================================================================
Private Function BlockFirstRow(ws As Worksheet, marker As String) As Long
    Dim r As Long, lastR As Long
    BlockFirstRow = 0
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastR
        If InStr(1, CStr(ws.Cells(r, 1).Value), marker, vbTextCompare) = 1 Then
            BlockFirstRow = r + 2
            Exit Function
        End If
    Next r
End Function


Private Function NumOf(v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = 0
End Function


Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
