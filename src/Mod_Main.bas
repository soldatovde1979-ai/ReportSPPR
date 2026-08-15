Attribute VB_Name = "Mod_Main"
Option Explicit

Public Sub RunSupportAnalytics()
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSource As Worksheet
    On Error Resume Next: Set wsSource = wb.Worksheets("Data")
    If wsSource Is Nothing Then Set wsSource = wb.ActiveSheet
    On Error GoTo ErrHandler
    Dim lastRow As Long: lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then MsgBox "Таблица с данными пуста!", vbExclamation: Exit Sub
    Application.StatusBar = "Чтение данных и расчет SLA...": Application.ScreenUpdating = False
    Dim tickets() As TicketData: ReDim tickets(1 To lastRow - 1)
    Dim validCount As Long: validCount = 0, r As Long
    Dim maxYear As Long: maxYear = 0, maxWeek As Long: maxWeek = 0
    Dim minRegDate As Date: minRegDate = DateSerial(2099, 1, 1), maxRegDate As Date: maxRegDate = DateSerial(1900, 1, 1)
    Dim minClsDate As Date: minClsDate = DateSerial(2099, 1, 1), maxClsDate As Date: maxClsDate = DateSerial(1900, 1, 1)
    For r = 2 To lastRow
        Dim clsVal As String: clsVal = Trim(CStr(wsSource.Cells(r, 3).Value))
        If clsVal <> "" And IsDate(clsVal) Then
            validCount = validCount + 1
            With tickets(validCount)
                .id = CStr(wsSource.Cells(r, 1).Value): .reg = CStr(wsSource.Cells(r, 2).Value): .cls = clsVal
                .cli = CStr(wsSource.Cells(r, 4).Value): .anl = CStr(wsSource.Cells(r, 5).Value)
                .sec = CStr(wsSource.Cells(r, 6).Value): .desc = CStr(wsSource.Cells(r, 7).Value): .sol = CStr(wsSource.Cells(r, 8).Value)
                .reg_dt = CDate(.reg): .cls_dt = CDate(.cls): .has_cls = True
                .iso_year = GetISOYear(.cls_dt): .iso_week = GetISOWeek(.cls_dt): .dif_Hour = CalcBusHours(.reg_dt, .cls_dt)
                If .iso_year > maxYear Then maxYear = .iso_year
                If .iso_week > maxWeek Then maxWeek = .iso_week
                If .reg_dt < minRegDate Then minRegDate = .reg_dt
                If .reg_dt > maxRegDate Then maxRegDate = .reg_dt
                If .cls_dt < minClsDate Then minClsDate = .cls_dt
                If .cls_dt > maxClsDate Then maxClsDate = .cls_dt
            End With
        End If
    Next r
    If validCount = 0 Then MsgBox "Нет закрытых тикетов!", vbExclamation: Exit Sub
    ReDim Preserve tickets(1 To validCount)
    Dim totalBatchTok As Long: totalBatchTok = 0, batchIdx As Long, numBatches As Long
    numBatches = Application.WorksheetFunction.RoundUp(validCount / BATCH_SIZE, 0)
    For batchIdx = 1 To numBatches
        Dim bStart As Long: bStart = (batchIdx - 1) * BATCH_SIZE + 1
        Dim bEnd As Long: bEnd = batchIdx * BATCH_SIZE: If bEnd > validCount Then bEnd = validCount
        Application.StatusBar = "Батч " & batchIdx & " из " & numBatches & "..."
        Dim promptRows As String: promptRows = "", i As Long
        For i = bStart To bEnd
            promptRows = promptRows & "ID: " & tickets(i).id & " | Тема: " & tickets(i).sec & " | Описание: " & tickets(i).desc & vbCrLf
        Next i
        Dim pTok As Long, cTok As Long, tTok As Long
        Dim jsonResp As String: jsonResp = CallGeminiBatch("Оцени:" & vbCrLf & promptRows, pTok, cTok, tTok)
        totalBatchTok = totalBatchTok + tTok
        Dim resObjects As Collection: Set resObjects = ExtractArrayObjects(jsonResp, "results")
        Dim itemJson As Variant
        For Each itemJson In resObjects
            Dim tId As String: tId = ExtractJSONValue(CStr(itemJson), "id")
            For i = bStart To bEnd
                If tickets(i).id = tId Then
                    tickets(i).topic_cluster = ExtractJSONValue(CStr(itemJson), "topic_cluster")
                    tickets(i).tone_deviation_score = CLng(Val(ExtractJSONValue(CStr(itemJson), "tone_deviation_score")))
                    tickets(i).complexity = CLng(Val(ExtractJSONValue(CStr(itemJson), "complexity")))
                    tickets(i).final_grade = UCase(Trim(ExtractJSONValue(CStr(itemJson), "final_grade")))
                    tickets(i).grade_argument = ExtractJSONValue(CStr(itemJson), "grade_argument")
                    Exit For
                End If
            Next i
        Next itemJson
    Next batchIdx
    Dim difArr() As Double: ReDim difArr(1 To validCount)
    Dim totalDifHours As Double: totalDifHours = 0#
    For i = 1 To validCount: difArr(i) = tickets(i).dif_Hour: totalDifHours = totalDifHours + tickets(i).dif_Hour: Next i
    Dim medHour As Double: medHour = Round(CalcMedian(difArr, validCount), 1)
    Dim midHour As Double: midHour = Round(totalDifHours / validCount, 1)
    Application.StatusBar = "Генерация выводов Pro..."
    Dim summaryTokens As Long: summaryTokens = 0
    Dim aiConclusions As String: aiConclusions = CallGeminiSummary("Всего: " & validCount & ", Медиана: " & medHour, summaryTokens)
    Dim runId As String: runId = Format(Now, "yyyymmdd_hhnnss")
    LogResults wb, runId, maxYear, maxWeek, validCount, totalBatchTok, summaryTokens
    Dim templatePath As String: templatePath = wb.Path & "\template.html"
    Dim outputPath As String: outputPath = wb.Path & "\Отчет по качеству W" & maxWeek & ".html"
    GenerateHtmlReport templatePath, outputPath, tickets, validCount, maxYear, maxWeek, _
                       Format(minRegDate, "dd.mm.yyyy"), Format(maxRegDate, "dd.mm.yyyy"), _
                       Format(minClsDate, "dd.mm.yyyy"), Format(maxClsDate, "dd.mm.yyyy"), _
                       medHour, midHour, aiConclusions, totalBatchTok, summaryTokens
    Application.StatusBar = False: Application.ScreenUpdating = True
    MsgBox "Готово! Отчет: " & outputPath, vbInformation
    Exit Sub
ErrHandler:
    Application.StatusBar = False: Application.ScreenUpdating = True
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Private Sub LogResults(ByVal wb As Workbook, ByVal runId As String, ByVal y As Long, ByVal w As Long, ByVal total As Long, ByVal bTok As Long, ByVal sTok As Long)
    Dim wsLogs As Worksheet, wsTokens As Worksheet
    On Error Resume Next: Set wsLogs = wb.Worksheets("Logs"): On Error GoTo 0
    If wsLogs Is Nothing Then
        Set wsLogs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): wsLogs.name = "Logs"
        wsLogs.Range("A1:F1").Value = Array("Run_ID", "Timestamp", "Year", "Week", "Total_Tickets", "Status")
    End If
    Dim nextR As Long: nextR = wsLogs.Cells(wsLogs.Rows.count, "A").End(xlUp).Row + 1
    wsLogs.Cells(nextR, 1).Resize(1, 6).Value = Array(runId, Format(Now, "yyyy-mm-dd hh:nn:ss"), y, w, total, "SUCCESS")
    On Error Resume Next: Set wsTokens = wb.Worksheets("Token"): On Error GoTo 0
    If wsTokens Is Nothing Then
        Set wsTokens = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): wsTokens.name = "Token"
        wsTokens.Range("A1:E1").Value = Array("Run_ID", "Batch_Model", "Batch_Tokens", "Summary_Model", "Summary_Tokens")
    End If
    nextR = wsTokens.Cells(wsTokens.Rows.count, "A").End(xlUp).Row + 1
    wsTokens.Cells(nextR, 1).Resize(1, 5).Value = Array(runId, BATCH_MODEL, bTok, SUMMARY_MODEL, sTok)
End Sub