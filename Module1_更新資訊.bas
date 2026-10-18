Attribute VB_Name = "Module1"
Option Explicit

'==================================================================================
'  基本補助費概算表 － 自動產製
'
'  法源：桃園市高級中等以下學校參加國內運動競賽補助要點「基本補助費基準表」
'
'  輸入：工作表「競賽資訊」（每列一場比賽）
'        工作表「費率對照」（三個區塊：地區費率 / 縣市對照 / 路線車資）
'  輸出：工作表「基本補助費」或「基本補助費-<場次名稱>」
'
'  注意：凡名稱以「基本補助費」開頭的工作表，每次執行時都會被刪除重建。
'        請勿把手改的資料放在這類工作表中。
'==================================================================================

Private Const SHEET_INFO As String = "競賽資訊"
Private Const SHEET_RATE As String = "費率對照"
Private Const OUT_PREFIX As String = "基本補助費"

Private Const MK_REGION As String = "【一】"
Private Const MK_CITY   As String = "【二】"
Private Const MK_FARE   As String = "【三】"

Private Const FALLBACK_FARE_KEY As String = "＊預設"

' 查不到地區時暫用的地區，以及自動新增列的待確認標記
Private Const DEFAULT_REGION As String = "中部、南部及東部地區"
Private Const PENDING_MARK   As String = "★待確認"


'----------------------------------------------------------------------------------
' 進入點
'----------------------------------------------------------------------------------
Public Sub 更新資訊()
    RunUpdate False
End Sub

' 供「競賽資訊」O8 的【計算差旅費】按鈕呼叫
Public Sub 計算差旅費()
    RunUpdate False
End Sub

' 供【一鍵計算差旅費.bat】呼叫：不跳任何對話框
Public Sub 更新資訊_靜默()
    RunUpdate True
End Sub


'----------------------------------------------------------------------------------
' 主流程
'----------------------------------------------------------------------------------
Private Sub RunUpdate(silent As Boolean)
    Dim wsInfo As Worksheet, wsRate As Worksheet
    Dim lastRow As Long, r As Long, i As Long, n As Long
    Dim hitRows() As Long
    Dim madeCount As Long
    Dim warnings As String, msg As String
    Dim singleMode As Boolean

    On Error GoTo ErrHandler

    Set wsInfo = ThisWorkbook.Worksheets(SHEET_INFO)
    Set wsRate = ThisWorkbook.Worksheets(SHEET_RATE)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    DeleteGeneratedSheets

    lastRow = wsInfo.Cells(wsInfo.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ReDim hitRows(1 To lastRow)
    n = 0
    For r = 2 To lastRow
        If IsCompetitionRow(wsInfo, r) Then
            n = n + 1
            hitRows(n) = r
        End If
    Next r

    If n = 0 Then
        If Not silent Then
            MsgBox "「競賽資訊」中找不到有效的比賽資料。" & vbLf & vbLf & _
                   "每一列至少需要：" & vbLf & _
                   "  A欄 比賽名稱、B欄 起始日期、C欄 結束日期、E欄 選手人數(>0)", _
                   vbExclamation, "沒有資料"
        End If
        GoTo CleanUp
    End If

    singleMode = (n = 1)

    For i = 1 To n
        If BuildOneSheet(wsInfo, wsRate, hitRows(i), singleMode, warnings) Then
            madeCount = madeCount + 1
        End If
    Next i

    RefreshPendingMarks wsRate
    wsInfo.Activate

    If Not silent Then
        msg = "已產製 " & madeCount & " 張「" & OUT_PREFIX & "」概算表。"
        If Len(warnings) > 0 Then
            msg = msg & vbLf & vbLf & "【請確認以下事項】" & vbLf & warnings
            MsgBox msg, vbExclamation, "完成（有提醒）"
        Else
            MsgBox msg, vbInformation, "完成"
        End If
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not silent Then
        MsgBox "發生錯誤：" & vbLf & Err.Number & " － " & Err.Description, vbCritical, "錯誤"
    End If
End Sub


'----------------------------------------------------------------------------------
' 產生單一場次的概算表
'----------------------------------------------------------------------------------
Private Function BuildOneSheet(wsInfo As Worksheet, wsRate As Worksheet, r As Long, _
                               singleMode As Boolean, ByRef warnings As String) As Boolean

    Dim ws As Worksheet
    Dim titleStr As String, locStr As String, srcStr As String
    Dim dStart As Date, dEnd As Date
    Dim days As Long, fundDays As Long, stayNights As Long
    Dim players As Long, staff As Long, headCount As Long
    Dim needPrep As Boolean, transMode As String
    Dim region As String, rr As Long
    Dim regionPending As Boolean, fareMissing As Boolean
    Dim mealRate As Double, roomRate As Double, fareMode As String
    Dim carCap As Double, commuteCap As Double
    Dim fare As Double, fareDetail As String, locExtra As String
    Dim fareOverride As Double
    Dim tripTotal As Long, carDays As Long
    Dim dailyFare As Double, capped As Boolean
    Dim transCell As String, transNote As String, transFormula As String
    Dim items As Double, perItem As Double, cleanFee As Double
    Dim rocYear As Long, noteN As Long
    Dim nm As String

    BuildOneSheet = False

    '--- 讀取輸入 -----------------------------------------------------------------
    titleStr = Trim$(CStr(wsInfo.Cells(r, 1).Value))
    dStart = CDate(wsInfo.Cells(r, 2).Value)
    dEnd = CDate(wsInfo.Cells(r, 3).Value)
    locStr = Trim$(CStr(wsInfo.Cells(r, 4).Value))
    players = CLng(NumOf(wsInfo.Cells(r, 5).Value))
    needPrep = (InStr(1, CStr(wsInfo.Cells(r, 7).Value), "是", vbTextCompare) > 0)
    transMode = Trim$(CStr(wsInfo.Cells(r, 8).Value))
    fareOverride = NumOf(wsInfo.Cells(r, 9).Value)
    items = NumOf(wsInfo.Cells(r, 10).Value)
    perItem = NumOf(wsInfo.Cells(r, 11).Value)
    cleanFee = NumOf(wsInfo.Cells(r, 12).Value)
    srcStr = Trim$(CStr(wsInfo.Cells(r, 13).Value))

    If srcStr = "" Then srcStr = "基本補助費"
    If transMode = "" Then transMode = "大眾運輸"

    If dEnd < dStart Then
        warnings = warnings & "‧「" & titleStr & "」結束日期早於起始日期，已略過。" & vbLf
        Exit Function
    End If

    '--- 日數（實際參賽日數，含頭尾） ---------------------------------------------
    days = DateDiff("d", dStart, dEnd) + 1

    '--- 補助人數 -----------------------------------------------------------------
    '    職員人數一律以「競賽資訊」F欄填列者為準（空白視為 0），不自動推算
    staff = CLng(NumOf(wsInfo.Cells(r, 6).Value))
    headCount = players + staff

    '--- 地區與費率 ---------------------------------------------------------------
    region = LookupRegion(wsRate, locStr, regionPending)
    If region = "" Then
        ' 查不到 → 自動新增到【二】，地區暫填 DEFAULT_REGION 並標記待確認
        region = DEFAULT_REGION
        AppendCityRow wsRate, locStr, region
        regionPending = True
        warnings = warnings & "‧已將比賽地點「" & locStr & "」自動新增至「費率對照」【二】，" & _
                   "地區暫定為「" & region & "」。請確認地區是否正確，改好後清除該列C欄的「" & _
                   PENDING_MARK & "」。" & vbLf
    ElseIf regionPending Then
        warnings = warnings & "‧「費率對照」【二】中「" & locStr & "」的地區仍標示「" & PENDING_MARK & _
                   "」，本次以「" & region & "」計算，請確認。" & vbLf
    End If

    rr = RegionRow(wsRate, region)
    If rr = 0 Then
        warnings = warnings & "‧「費率對照」【一】找不到地區「" & region & "」，已略過「" & titleStr & "」。" & vbLf
        Exit Function
    End If
    mealRate = NumOf(wsRate.Cells(rr, 2).Value)
    roomRate = NumOf(wsRate.Cells(rr, 3).Value)
    fareMode = Trim$(CStr(wsRate.Cells(rr, 4).Value))
    carCap = NumOf(wsRate.Cells(rr, 5).Value)
    commuteCap = NumOf(wsRate.Cells(rr, 6).Value)

    '--- 補助日數：基準表備註（四） -----------------------------------------------
    '    以實際參賽日數審核；外島／中南東部需事前移地準備者得加一日
    fundDays = days
    If needPrep Then
        If roomRate > 0 And transMode <> "每日通勤" Then
            fundDays = days + 1
        Else
            warnings = warnings & "‧「" & titleStr & "」屬" & region & _
                       "（或採每日通勤），依基準表備註（四）不適用「移地準備加一日」，已忽略。" & vbLf
        End If
    End If

    '--- 住宿晚數：末日屬折返日，不核予住宿費 -------------------------------------
    If roomRate > 0 And transMode <> "每日通勤" Then
        stayNights = fundDays - 1
        If stayNights < 0 Then stayNights = 0
    Else
        stayNights = 0
    End If

    '--- 車資 ---------------------------------------------------------------------
    fare = LookupFare(wsRate, locStr, fareDetail, locExtra, fareMissing)
    If fareOverride > 0 Then
        ' 競賽資訊 I 欄已人工填列，直接採用，不動【三】
        fare = fareOverride
        fareDetail = "單趟票價 " & Format$(fare, "#,##0") & " 元（人工填列）"
    ElseIf fareMissing And transMode <> "租賃車輛" Then
        ' 查不到路線 → 自動新增到【三】（插在「＊預設」之前）
        AppendFareRow wsRate, locStr
        fareDetail = "（待填票價）"
        warnings = warnings & "‧已將比賽地點「" & locStr & "」自動新增至「費率對照」【三】路線車資表，" & _
                   "請填入單趟票價與路線明細後重算（本次交通費以 0 計）。" & vbLf
    ElseIf fare <= 0 And transMode <> "租賃車輛" Then
        warnings = warnings & "‧「費率對照」【三】中「" & locStr & "」的單趟票價仍為 0，" & _
                   "請填入實際票價後重算（本次交通費以 0 計）。" & vbLf
        If fareDetail = "" Then fareDetail = "（待填票價）"
    End If

    '--- 交通費：基準表交通費欄位 + 備註（二）（三） -------------------------------
    Select Case transMode

        Case "租賃車輛"
            ' 備註（三）：桃園、北部以補助一日為限；外島、中南東部因賽程有住宿需求者最多二日
            If roomRate > 0 And stayNights > 0 Then carDays = 2 Else carDays = 1
            transFormula = "=" & carCap & "*" & carDays
            transCell = "租賃車輛（不含一般計程車及多元化計程車）" & vbLf & _
                        "(" & region & ")每日上限 " & Format$(carCap, "#,##0") & " 元" & vbLf & _
                        carDays & "日 × " & Format$(carCap, "#,##0") & "元（以上限估列，核實補助）"
            transNote = "採租賃車輛，" & region & "每日上限" & Format$(carCap, "#,##0") & "元；" & _
                        "依基準表備註（三）補助" & carDays & "日，以上限估列，核實補助。"

        Case "每日通勤"
            ' 備註（二）：中南東部不於當地住宿採每日通勤者，每日申請金額不得逾 700 元
            dailyFare = fare * 2
            capped = False
            If commuteCap > 0 And dailyFare > commuteCap Then
                dailyFare = commuteCap
                capped = True
                warnings = warnings & "‧「" & titleStr & "」採每日通勤，往返 " & Format$(fare * 2, "#,##0") & _
                           " 元已逾基準表備註（二）每日 " & Format$(commuteCap, "#,##0") & _
                           " 元上限，已依上限計列。" & vbLf
            End If
            transFormula = "=" & headCount & "*" & fundDays & "*" & dailyFare
            transCell = fareDetail & vbLf & _
                        "採每日通勤往返（不列住宿費）" & vbLf & _
                        headCount & "人 × " & fundDays & "日 × 每人每日" & Format$(dailyFare, "#,##0") & "元" & _
                        IIf(capped, vbLf & "＊已依基準表備註（二）每日" & Format$(commuteCap, "#,##0") & "元上限計列", "")
            transNote = "採每日通勤往返，" & OneLine(fareDetail) & "，每人每日" & Format$(dailyFare, "#,##0") & "元；" & _
                        "計 " & headCount & "人×" & fundDays & "日" & _
                        IIf(capped, "（已套用每日" & Format$(commuteCap, "#,##0") & "元上限）", "") & "。"

        Case Else   ' 大眾運輸
            transMode = "大眾運輸"
            If InStr(fareMode, "每日") > 0 Then
                ' 桃園市、北部地區：依實際路程核給「每日往返一次」
                tripTotal = fundDays * 2
                transCell = fareDetail & vbLf & _
                            headCount & "人 × " & fundDays & "日 × 2趟 × 每趟" & Format$(fare, "#,##0") & "元" & vbLf & _
                            "＊(" & region & ")依基準表核給「每日往返一次」"
                transNote = OneLine(fareDetail) & "＝" & Format$(fare, "#,##0") & "元/一趟；" & _
                            region & "依基準表核給每日往返一次，計 " & headCount & "人×" & fundDays & "日×2趟。"
            Else
                ' 中南東部、外島：依實際路程核給「往返一次」
                tripTotal = 2
                transCell = fareDetail & vbLf & _
                            headCount & "人 × 2趟 × 每趟" & Format$(fare, "#,##0") & "元" & vbLf & _
                            "＊(" & region & ")依基準表核給「往返一次」"
                transNote = OneLine(fareDetail) & "＝" & Format$(fare, "#,##0") & "元/一趟；" & _
                            region & "依基準表核給往返一次，計 " & headCount & "人×2趟。"
            End If
            transFormula = "=" & fare & "*" & tripTotal & "*" & headCount
    End Select

    '--- 建立工作表 ---------------------------------------------------------------
    If singleMode Then
        nm = UniqueSheetName(OUT_PREFIX)
    Else
        nm = UniqueSheetName(OUT_PREFIX & "-" & titleStr)
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm

    With ws
        .Range("A1").Value = titleStr & "　概算表"

        .Range("A2").Value = "項次"
        .Range("B2").Value = "項目"
        .Range("C2").Value = "明細"
        .Range("D2").Value = "預估經費"
        .Range("E2").Value = "經費來源"

        ' 1 報名費（核實補助；僅選手負擔報名費）
        .Range("A3").Value = 1
        .Range("B3").Value = "報名費"
        .Range("C3").Value = "核實補助（依大會簡章）" & vbLf & _
                             "選手" & players & "人 × (" & items & "項 × 每項" & Format$(perItem, "#,##0") & "元" & _
                             " + 每人清潔費" & Format$(cleanFee, "#,##0") & "元)"
        .Range("D3").Formula = "=" & players & "*(" & items & "*" & perItem & "+" & cleanFee & ")"
        .Range("E3").Value = srcStr

        ' 2 住宿費
        .Range("A4").Value = 2
        .Range("B4").Value = "住宿費"
        If stayNights > 0 Then
            .Range("C4").Value = "(" & region & ")每人每晚" & Format$(roomRate, "#,##0") & "元" & vbLf & _
                                 headCount & "人 × " & stayNights & "晚 × " & Format$(roomRate, "#,##0") & "元" & vbLf & _
                                 "＊補助日數" & fundDays & "日" & IIf(fundDays > days, "（含移地準備1日）", "") & _
                                 "，末日屬折返日不列住宿，故提列" & stayNights & "晚"
            .Range("D4").Formula = "=" & headCount & "*" & stayNights & "*" & roomRate
        Else
            If roomRate = 0 Then
                .Range("C4").Value = "(" & region & ")依基準表不核予住宿費"
            ElseIf transMode = "每日通勤" Then
                .Range("C4").Value = "採每日通勤往返，不列住宿費"
            Else
                .Range("C4").Value = "實際參賽" & days & "日，末日屬折返日" & vbLf & "依基準表備註（四）不核予住宿費"
            End If
            .Range("D4").Formula = "=0"
        End If

        ' 3 膳費
        .Range("A5").Value = 3
        .Range("B5").Value = "膳費"
        .Range("C5").Value = "(" & region & ")每人每日" & Format$(mealRate, "#,##0") & "元" & vbLf & _
                             headCount & "人 × " & fundDays & "日 × " & Format$(mealRate, "#,##0") & "元" & _
                             IIf(fundDays > days, vbLf & "＊需事前移地準備，依基準表備註（四）加計1日", "")
        .Range("D5").Formula = "=" & headCount & "*" & fundDays & "*" & mealRate

        ' 4 交通費
        .Range("A6").Value = 4
        .Range("B6").Value = "交通費"
        .Range("C6").Value = transCell
        .Range("D6").Formula = transFormula

        ' 合計
        .Range("A7").Value = "合計"
        .Range("D7").Formula = "=SUM(D3:D6)"

        ' 備註
        .Range("A8").Value = "備註："
        rocYear = Year(dStart) - 1911
        noteN = 9
        .Cells(noteN, 1).Value = "1.比賽日期：" & rocYear & "年" & Format$(dStart, "m/d") & "－" & _
                                 Format$(dEnd, "m/d") & "，實際參賽" & days & "日" & _
                                 IIf(fundDays > days, "（需事前移地準備，補助日數" & fundDays & "日）", "") & "。"
        noteN = noteN + 1
        .Cells(noteN, 1).Value = "2.比賽地點：" & locStr & IIf(locExtra <> "", "－" & locExtra, "") & _
                                 "（依基準表屬" & region & "）"
        noteN = noteN + 1
        .Cells(noteN, 1).Value = "3.補助人數：選手" & players & "名、職員" & staff & "名，共" & headCount & "名。"
        noteN = noteN + 1
        .Cells(noteN, 1).Value = "4.交通費：" & transNote
        noteN = noteN + 1
        If roomRate = 0 Then
            .Cells(noteN, 1).Value = "5.住宿費：" & region & "依基準表不核予住宿費。"
        ElseIf transMode = "每日通勤" Then
            .Cells(noteN, 1).Value = "5.住宿費：採每日通勤往返，不列住宿費；交通費依基準表備註（二）以每人每日" & _
                                     Format$(commuteCap, "#,##0") & "元為上限。"
        Else
            .Cells(noteN, 1).Value = "5.住宿費：依基準表備註（四），參賽最後一日屬折返日，不核予住宿費。"
        End If
        noteN = noteN + 1
        .Cells(noteN, 1).Value = "6.本表依「桃園市高級中等以下學校參加國內運動競賽補助要點－基本補助費基準表」計列。"

        '--- 格式 -----------------------------------------------------------------
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 48
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 12

        .Range("A1:E1").Merge
        With .Range("A1")
            .Font.Size = 14
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(1).RowHeight = 28

        With .Range("A2:E2")
            .Font.Bold = True
            .Interior.Color = RGB(217, 234, 247)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With .Range("A2:E7")
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(128, 128, 128)
            .VerticalAlignment = xlCenter
        End With

        .Range("C3:C6").WrapText = True
        .Range("A3:B7").HorizontalAlignment = xlCenter
        .Range("D3:D7").NumberFormat = "#,##0"
        .Range("E3:E7").HorizontalAlignment = xlCenter

        .Range("A7:C7").Merge
        With .Range("A7")
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        With .Range("A7:E7")
            .Font.Bold = True
            .Interior.Color = RGB(255, 242, 204)
        End With

        .Range("A8").Font.Bold = True
        .Rows("3:6").EntireRow.AutoFit
        .Range("A1").Select
    End With

    BuildOneSheet = True
End Function


'==================================================================================
'  工具函式
'==================================================================================

' 判定某列是否為有效的比賽資料（可靠地排除說明文字列）
Private Function IsCompetitionRow(ws As Worksheet, r As Long) As Boolean
    IsCompetitionRow = False
    If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then Exit Function
    If Not IsDate(ws.Cells(r, 2).Value) Then Exit Function
    If Not IsDate(ws.Cells(r, 3).Value) Then Exit Function
    If NumOf(ws.Cells(r, 5).Value) <= 0 Then Exit Function
    IsCompetitionRow = True
End Function


Private Function NumOf(v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = 0
End Function


Private Function OneLine(s As String) As String
    OneLine = Replace(Replace(s, vbCrLf, "＋"), vbLf, "＋")
End Function


' 找出「費率對照」中某區塊的第一列資料（標記列 + 欄位標題列的下一列）
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


' 依比賽地點文字比對縣市關鍵字，回傳基準表地區名稱
' pending：命中的那一列 C 欄有「★待確認」標記（表示是自動新增、地區尚未人工確認）
Private Function LookupRegion(wsRate As Worksheet, locStr As String, ByRef pending As Boolean) As String
    Dim r As Long, kw As String
    LookupRegion = ""
    pending = False
    r = BlockFirstRow(wsRate, MK_CITY)
    If r = 0 Then Exit Function
    Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
        kw = Trim$(CStr(wsRate.Cells(r, 1).Value))
        If InStr(1, locStr, kw, vbTextCompare) > 0 Then
            LookupRegion = Trim$(CStr(wsRate.Cells(r, 2).Value))
            pending = (Trim$(CStr(wsRate.Cells(r, 3).Value)) <> "")
            Exit Function
        End If
        r = r + 1
    Loop
End Function


' 取得地區在【一】地區費率表中的列號
Private Function RegionRow(wsRate As Worksheet, region As String) As Long
    Dim r As Long
    RegionRow = 0
    r = BlockFirstRow(wsRate, MK_REGION)
    If r = 0 Then Exit Function
    Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
        If StrComp(Trim$(CStr(wsRate.Cells(r, 1).Value)), region, vbTextCompare) = 0 Then
            RegionRow = r
            Exit Function
        End If
        r = r + 1
    Loop
End Function


' 依比賽地點查【三】路線車資表，回傳單趟票價
' hitFallback：沒有任何路線命中，只落到「＊預設」列（代表這個地點還沒建路線）
Private Function LookupFare(wsRate As Worksheet, locStr As String, _
                            ByRef fareDetail As String, ByRef locExtra As String, _
                            ByRef hitFallback As Boolean) As Double
    Dim r As Long, kw As String
    LookupFare = 0
    fareDetail = ""
    locExtra = ""
    hitFallback = True          ' 連【三】都找不到時，同樣視為需要新增
    r = BlockFirstRow(wsRate, MK_FARE)
    If r = 0 Then Exit Function
    Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
        kw = Trim$(CStr(wsRate.Cells(r, 1).Value))
        If kw = FALLBACK_FARE_KEY Or InStr(1, locStr, kw, vbTextCompare) > 0 Then
            LookupFare = NumOf(wsRate.Cells(r, 2).Value)
            fareDetail = CStr(wsRate.Cells(r, 3).Value)
            locExtra = Trim$(CStr(wsRate.Cells(r, 4).Value))
            hitFallback = (kw = FALLBACK_FARE_KEY)
            Exit Function
        End If
        r = r + 1
    Loop
End Function


'----------------------------------------------------------------------------------
' 自動把查不到的比賽地點寫回「費率對照」
'----------------------------------------------------------------------------------

' 新增一列到【二】縣市→地區對照表的最末端
' （放最後才不會蓋掉既有的縣市關鍵字；先插入空列再寫，以保留區塊後的空白分隔列）
Private Sub AppendCityRow(wsRate As Worksheet, kw As String, region As String)
    Dim r As Long
    r = BlockFirstRow(wsRate, MK_CITY)
    If r = 0 Then Exit Sub
    Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
        r = r + 1
    Loop
    wsRate.Rows(r).Insert Shift:=xlDown
    wsRate.Cells(r, 1).Value = kw
    wsRate.Cells(r, 2).Value = region
    wsRate.Cells(r, 3).Value = PENDING_MARK
    HighlightRow wsRate, r, 3
End Sub


' 新增一列到【三】路線車資表，插在「＊預設」列之前（＊預設必須永遠留在最後）
Private Sub AppendFareRow(wsRate As Worksheet, kw As String)
    Dim r As Long
    r = BlockFirstRow(wsRate, MK_FARE)
    If r = 0 Then Exit Sub
    Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
        If Trim$(CStr(wsRate.Cells(r, 1).Value)) = FALLBACK_FARE_KEY Then Exit Do
        r = r + 1
    Loop
    wsRate.Rows(r).Insert Shift:=xlDown
    wsRate.Cells(r, 1).Value = kw
    wsRate.Cells(r, 2).Value = 0
    wsRate.Cells(r, 3).Value = "（待填：請填入實際單趟票價與路線明細，並檢附票價圖）"
    wsRate.Cells(r, 4).Value = ""
    HighlightRow wsRate, r, 4
End Sub


Private Sub HighlightRow(wsRate As Worksheet, r As Long, nCols As Long)
    With wsRate.Range(wsRate.Cells(r, 1), wsRate.Cells(r, nCols))
        .Interior.Color = RGB(255, 242, 153)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With
    wsRate.Rows(r).EntireRow.AutoFit
End Sub


' 使用者補齊資料後，把不再待確認的列黃底清掉
Private Sub RefreshPendingMarks(wsRate As Worksheet)
    Dim r As Long
    r = BlockFirstRow(wsRate, MK_CITY)
    If r > 0 Then
        Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
            If Trim$(CStr(wsRate.Cells(r, 3).Value)) = "" Then
                wsRate.Range(wsRate.Cells(r, 1), wsRate.Cells(r, 3)).Interior.Pattern = xlNone
            End If
            r = r + 1
        Loop
    End If
    r = BlockFirstRow(wsRate, MK_FARE)
    If r > 0 Then
        Do While Trim$(CStr(wsRate.Cells(r, 1).Value)) <> ""
            If NumOf(wsRate.Cells(r, 2).Value) > 0 Then
                wsRate.Range(wsRate.Cells(r, 1), wsRate.Cells(r, 4)).Interior.Pattern = xlNone
            End If
            r = r + 1
        Loop
    End If
End Sub


' 刪除上次產製的所有概算表（名稱以「基本補助費」開頭者）
Private Sub DeleteGeneratedSheets()
    Dim i As Long, ws As Worksheet
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left$(ws.Name, Len(OUT_PREFIX)) = OUT_PREFIX Then
            If ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
        End If
    Next i
End Sub


' 工作表名稱合法化（去除 : \ / ? * [ ] 並限制 31 字）
Private Function SafeSheetName(baseName As String) As String
    Dim s As String, bad As Variant, i As Long
    s = Trim$(baseName)
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    If Len(s) > 31 Then s = Left$(s, 31)
    If s = "" Then s = OUT_PREFIX
    SafeSheetName = s
End Function


' 名稱重複時自動加序號
Private Function UniqueSheetName(baseName As String) As String
    Dim s As String, cand As String, n As Long
    s = SafeSheetName(baseName)
    cand = s
    n = 1
    Do While SheetExists(cand)
        n = n + 1
        cand = SafeSheetName(Left$(s, 27) & "(" & n & ")")
    Loop
    UniqueSheetName = cand
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
