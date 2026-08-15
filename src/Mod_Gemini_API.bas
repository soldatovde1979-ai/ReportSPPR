Attribute VB_Name = "Mod_Gemini_API"
Option Explicit

Public Function CallGeminiBatch(ByVal promptText As String, ByRef outPromptTok As Long, ByRef outCandTok As Long, ByRef outTotalTok As Long) As String
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim url As String: url = "https://generativelanguage.googleapis.com/v1beta/models/" & BATCH_MODEL & ":generateContent?key=" & GEMINI_API_KEY
    Dim schemaJson As String
    schemaJson = "{""type"":""OBJECT"",""properties"":{""results"":{""type"":""ARRAY"",""items"":{""type"":""OBJECT"",""properties"":{" & _
                 """id"":{""type"":""STRING""},""topic_cluster"":{""type"":""STRING""},""tone_deviation_score"":{""type"":""INTEGER""}," & _
                 """is_emotional"":{""type"":""BOOLEAN""},""emotion_argument"":{""type"":""STRING""},""complexity"":{""type"":""INTEGER""}," & _
                 """initial_grade"":{""type"":""STRING""},""final_grade"":{""type"":""STRING""},""grade_argument"":{""type"":""STRING""}," & _
                 """is_alternative"":{""type"":""BOOLEAN""},""is_critical_incident"":{""type"":""BOOLEAN""}" & _
                 "},""required"":[""id"",""topic_cluster"",""tone_deviation_score"",""is_emotional"",""emotion_argument"",""complexity"",""initial_grade"",""final_grade"",""grade_argument"",""is_alternative"",""is_critical_incident""]}}}}"
    Dim body As String
    body = "{""contents"":[{""parts"":[{""text"":""" & EscapeJSON(promptText) & """}]}]," & _
           """systemInstruction"":{""parts"":[{""text"":""" & EscapeJSON(GetSystemInstruction()) & """}]}]," & _
           """generationConfig"":{""responseMimeType"":""application/json"",""responseSchema"":" & schemaJson & ",""temperature"":0.1}}"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body
    If http.Status = 200 Then
        Dim respText As String: respText = http.responseText
        outPromptTok = CLng(Val(ExtractJSONValue(respText, "promptTokenCount")))
        outCandTok = CLng(Val(ExtractJSONValue(respText, "candidatesTokenCount")))
        outTotalTok = CLng(Val(ExtractJSONValue(respText, "totalTokenCount")))
        CallGeminiBatch = ExtractJSONValue(respText, "text")
    Else
        Err.Raise vbObjectError + 500, "CallGeminiBatch", "API Error " & http.Status & ": " & http.responseText
    End If
End Function

Public Function CallGeminiSummary(ByVal summaryPrompt As String, ByRef outTotalTok As Long) As String
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim url As String: url = "https://generativelanguage.googleapis.com/v1beta/models/" & SUMMARY_MODEL & ":generateContent?key=" & GEMINI_API_KEY
    Dim fullPrompt As String
    fullPrompt = "Ты — главный архитектор систем 1С. Дай 3 вывода Root Cause: " & summaryPrompt & ". Формат HTML <li>"
    Dim body As String
    body = "{""contents"":[{""parts"":[{""text"":""" & EscapeJSON(fullPrompt) & """}]}],""generationConfig"":{""temperature"":0.2}}"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body
    If http.Status = 200 Then
        Dim respText As String: respText = http.responseText
        outTotalTok = CLng(Val(ExtractJSONValue(respText, "totalTokenCount")))
        Dim res As String: res = ExtractJSONValue(respText, "text")
        res = Replace(Replace(res, "```html", ""), "```", "")
        CallGeminiSummary = Trim(res)
    Else
        CallGeminiSummary = "<li>Ошибка генерации выводов: " & http.Status & "</li>"
    End If
End Function