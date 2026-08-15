Attribute VB_Name = "Mod_Report_Generator"
Option Explicit

Public Sub GenerateHtmlReport( _
    ByVal templatePath As String, ByVal outputPath As String, _
    ByRef tickets() As TicketData, ByVal totalCount As Long, _
    ByVal maxYear As Long, ByVal maxWeek As Long, _
    ByVal minRegStr As String, ByVal maxRegStr As String, _
    ByVal minClsStr As String, ByVal maxClsStr As String, _
    ByVal medHour As Double, ByVal midHour As Double, _
    ByVal aiConclusions As String, _
    ByVal batchTokens As Long, ByVal summaryTokens As Long _
)
    Dim fso As Object, stm As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(templatePath) Then Err.Raise vbObjectError + 404, "ReportGen", "Шаблон не найден: " & templatePath
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.LoadFromFile templatePath
    Dim tpl As String: tpl = stm.ReadText(-1): stm.Close
    Dim emoTotal As Long, emoSum As Long, emoMax As Long, altTotal As Long, i As Long
    Dim topicCounts As Object: Set topicCounts = CreateObject("Scripting.Dictionary")
    Dim initiatorCounts As Object: Set initiatorCounts = CreateObject("Scripting.Dictionary")
    Dim initiatorTopics As Object: Set initiatorTopics = CreateObject("Scripting.Dictionary")
    Dim anlTotal As Object: Set anlTotal = CreateObject("Scripting.Dictionary")
    Dim anlExc As Object: Set anlExc = CreateObject("Scripting.Dictionary")
    Dim anlGood As Object: Set anlGood = CreateObject("Scripting.Dictionary")
    Dim anlBad As Object: Set anlBad = CreateObject("Scripting.Dictionary")
    Dim anlCompSum As Object: Set anlCompSum = CreateObject("Scripting.Dictionary")
    For i = 1 To totalCount
        If tickets(i).tone_deviation_score > 0 Then
            emoTotal = emoTotal + 1: emoSum = emoSum + tickets(i).tone_deviation_score
            If tickets(i).tone_deviation_score > emoMax Then emoMax = tickets(i).tone_deviation_score
        End If
        If tickets(i).is_alternative Then altTotal = altTotal + 1
        Dim cl As String: cl = tickets(i).topic_cluster
        If cl <> "" Then topicCounts(cl) = topicCounts(cl) + 1
        Dim cli As String: cli = tickets(i).cli
        If cli <> "" Then initiatorCounts(cli) = initiatorCounts(cli) + 1: initiatorTopics(cli) = initiatorTopics(cli) & IIf(initiatorTopics(cli) = "", "", ", ") & cl
        Dim anl As String: anl = tickets(i).anl
        If anl <> "" Then
            anlTotal(anl) = anlTotal(anl) + 1: anlCompSum(anl) = anlCompSum(anl) + tickets(i).complexity
            If tickets(i).final_grade = "ОТЛИЧНО" Then anlExc(anl) = anlExc(anl) + 1
            If tickets(i).final_grade = "ХОРОШО" Then anlGood(anl) = anlGood(anl) + 1
            If tickets(i).final_grade = "ПЛОХО" Then anlBad(anl) = anlBad(anl) + 1
        End If
    Next i
    Dim emoAvg As Double: emoAvg = IIf(emoTotal > 0, Round(CDbl(emoSum) / emoTotal, 1), 0#)
    Dim topInitHtml As String: topInitHtml = ""
    Dim kCli As Variant, initRows As Long: initRows = 0
    For Each kCli In initiatorCounts.Keys
        initRows = initRows + 1: If initRows > TOP_INITIATORS_N Then Exit For
        topInitHtml = topInitHtml & "<tr><td><b>" & kCli & "</b></td><td>" & initiatorCounts(kCli) & "</td><td>" & initiatorTopics(kCli) & "</td><td>-</td></tr>"
    Next kCli
    Dim topAnlHtml As String: topAnlHtml = ""
    Dim nomHardAnl As String, nomQtyAnl As String, nomQualAnl As String
    Dim maxScore As Double: maxScore = -9999, maxQty As Long: maxQty = 0
    For Each kCli In anlTotal.Keys
        Dim cTot As Long: cTot = anlTotal(kCli)
        Dim cExc As Long: cExc = anlExc(kCli)
        Dim cGood As Long: cGood = anlGood(kCli)
        Dim cBad As Long: cBad = anlBad(kCli)
        Dim avgCmp As Double: avgCmp = Round(CDbl(anlCompSum(kCli)) / cTot, 1)
        Dim sc As Double: sc = Round((cExc * 3 + cGood * 1 - cBad * 3) * avgCmp, 1)
        If cTot > maxQty Then maxQty = cTot: nomQtyAnl = CStr(kCli)
        If sc > maxScore Then maxScore = sc: nomQualAnl = CStr(kCli)
        topAnlHtml = topAnlHtml & "<tr><td><b>" & kCli & "</b></td><td>" & cTot & "</td><td>" & cExc & "</td><td>" & cGood & "</td><td>" & cBad & "</td><td>" & avgCmp & "</td><td><b>" & sc & "</b></td></tr>"
    Next kCli
    Dim pieJson As String: pieJson = "{"
    Dim firstK As Boolean: firstK = True
    For Each kCli In topicCounts.Keys
        If Not firstK Then pieJson = pieJson & ","
        pieJson = pieJson & """" & EscapeJSON(CStr(kCli)) & """:" & topicCounts(kCli): firstK = False
    Next kCli
    pieJson = pieJson & "}"
    Dim h As String: h = tpl
    h = Replace(h, "{{ week_header }}", "W" & maxWeek & " [" & minClsStr & "]")
    h = Replace(h, "{{ total_tickets }}", CStr(totalCount))
    h = Replace(h, "{{ med_hour }}", Format(medHour, "0.0"))
    h = Replace(h, "{{ mid_hour }}", Format(midHour, "0.0"))
    h = Replace(h, "{{ dates_reg }}", minRegStr & " — " & maxRegStr)
    h = Replace(h, "{{ dates_cls }}", minClsStr & " — " & maxClsStr)
    h = Replace(h, "{{ emo_total }}", CStr(emoTotal))
    h = Replace(h, "{{ emo_avg }}", Format(emoAvg, "0.0"))
    h = Replace(h, "{{ emo_max }}", CStr(emoMax))
    h = Replace(h, "{{ alt_total }}", CStr(altTotal))
    h = Replace(h, "{{ ai_conclusions | safe }}", aiConclusions)
    h = Replace(h, "{{ tip_text }}", "Оптимизация 3 частых сбоев снижает нагрузку на 15%.")
    h = Replace(h, "{{ pie_json | safe }}", pieJson)
    h = Replace(h, "{{ batch_tokens }}", CStr(batchTokens))
    h = Replace(h, "{{ summary_tokens }}", CStr(summaryTokens))
    h = Replace(h, "{{ batch_model }}", BATCH_MODEL)
    h = Replace(h, "{{ summary_model }}", SUMMARY_MODEL)
    h = Replace(h, "{{ nom_hard_anl }}", IIf(nomHardAnl = "", "Н/Д", nomHardAnl))
    h = Replace(h, "{{ nom_hard_txt }}", "Макс. сложность")
    h = Replace(h, "{{ nom_qty_anl }}", IIf(nomQtyAnl = "", "Н/Д", nomQtyAnl))
    h = Replace(h, "{{ nom_qty_txt }}", maxQty & " обращений")
    h = Replace(h, "{{ nom_qual_anl }}", IIf(nomQualAnl = "", "Н/Д", nomQualAnl))
    h = Replace(h, "{{ nom_qual_txt }}", "Балл: " & Format(maxScore, "0.0"))
    h = ReplaceJinjaBlock(h, "{% for row in top_initiators %}", "{% endfor %}", topInitHtml)
    h = ReplaceJinjaBlock(h, "{% for row in top_analysts %}", "{% endfor %}", topAnlHtml)
    Set stm = CreateObject("ADODB.Stream"): stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.WriteText h: stm.SaveToFile outputPath, 2: stm.Close
End Sub

Private Function ReplaceJinjaBlock(ByVal content As String, ByVal startTag As String, ByVal endTag As String, ByVal replacement As String) As String
    Dim p1 As Long, p2 As Long: p1 = InStr(content, startTag)
    If p1 > 0 Then
        p2 = InStr(p1, content, endTag)
        If p2 > 0 Then ReplaceJinjaBlock = Left(content, p1 - 1) & replacement & Mid(content, p2 + Len(endTag)): Exit Function
    End If
    ReplaceJinjaBlock = content
End Function