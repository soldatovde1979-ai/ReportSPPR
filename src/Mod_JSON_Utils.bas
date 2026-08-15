Attribute VB_Name = "Mod_JSON_Utils"
Option Explicit

Public Function EscapeJSON(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n")
    txt = Replace(txt, vbCr, "\n")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbTab, "\t")
    EscapeJSON = txt
End Function

Public Function ExtractJSONValue(ByVal json As String, ByVal key As String) As String
    Dim reg As Object, matches As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = False
    reg.IgnoreCase = True
    reg.Pattern = """" & key & """\s*:\s*(""([^""\\]*(\\.[^""\\]*)*)""|([0-9\.\-]+)|(true|false|null))"
    If reg.Test(json) Then
        Set matches = reg.Execute(json)
        Dim val As String: val = matches(0).SubMatches(0)
        If Left(val, 1) = """" And Right(val, 1) = """" Then
            val = Mid(val, 2, Len(val) - 2)
            val = Replace(val, "\""", """")
            val = Replace(val, "\n", vbCrLf)
            val = Replace(val, "\\", "\")
        End If
        ExtractJSONValue = val
    Else
        ExtractJSONValue = ""
    End If
End Function

Public Function ExtractArrayObjects(ByVal json As String, ByVal arrayName As String) As Collection
    Dim col As New Collection, pStart As Long
    pStart = InStr(1, json, """" & arrayName & """", vbTextCompare)
    If pStart = 0 Then Set ExtractArrayObjects = col: Exit Function
    pStart = InStr(pStart, json, "[")
    If pStart = 0 Then Exit Function
    Dim i As Long, depth As Long, objStart As Long, inQuotes As Boolean, ch As String
    depth = 0: objStart = 0: inQuotes = False
    For i = pStart To Len(json)
        ch = Mid(json, i, 1)
        If ch = """" And Mid(json, i - 1, 1) <> "\" Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "{" Then
                If depth = 0 Then objStart = i
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 And objStart > 0 Then
                    col.Add Mid(json, objStart, i - objStart + 1)
                    objStart = 0
                End If
            ElseIf ch = "]" And depth = 0 Then
                Exit For
            End If
        End If
    Next i
    Set ExtractArrayObjects = col
End Function