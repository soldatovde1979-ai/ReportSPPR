Attribute VB_Name = "Mod_Analytics"
Option Explicit

Public Function GetISOWeek(ByVal d As Date) As Long
    GetISOWeek = DatePart("ww", d, vbMonday, vbFirstFourDays)
End Function

Public Function GetISOYear(ByVal d As Date) As Long
    Dim w As Long, y As Long: w = DatePart("ww", d, vbMonday, vbFirstFourDays): y = Year(d)
    If Month(d) = 1 And w >= 52 Then GetISOYear = y - 1 ElseIf Month(d) = 12 And w = 1 Then GetISOYear = y + 1 Else GetISOYear = y
End Function

Public Function CalcBusHours(ByVal dtStart As Date, ByVal dtEnd As Date) As Double
    Dim d1 As Date, d2 As Date: d1 = DateValue(dtStart): d2 = DateValue(dtEnd)
    If d1 > d2 Then CalcBusHours = 0#: Exit Function
    Dim busDays As Long: busDays = Application.WorksheetFunction.NetworkDays(d1, d2) - 1
    If busDays < 0 Then busDays = 0
    CalcBusHours = busDays * WORK_HOURS
End Function

Public Function CalcMedian(ByRef arr() As Double, ByVal count As Long) As Double
    If count = 0 Then CalcMedian = 0#: Exit Function
    Dim i As Long, j As Long, temp As Double
    For i = 1 To count - 1
        For j = i + 1 To count
            If arr(i) > arr(j) Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
        Next j
    Next i
    If count Mod 2 = 1 Then CalcMedian = arr((count + 1) \ 2) Else CalcMedian = (arr(count \ 2) + arr((count \ 2) + 1)) / 2#
End Function