Attribute VB_Name = "Mod_Config_And_Types"
Option Explicit

Public Const GEMINI_API_KEY As String = "ВАШ_GEMINI_API_KEY"
Public Const BATCH_MODEL As String = "gemini-2.5-flash"
Public Const SUMMARY_MODEL As String = "gemini-2.5-pro"
Public Const BATCH_SIZE As Long = 40
Public Const WORK_HOURS As Double = 8#

Public Const TOP_INITIATORS_N As Long = 5
Public Const TOP_LONGEST_N As Long = 10
Public Const TOP_PROBLEMS_N As Long = 5

Public Type TicketData
    id As String
    reg As String
    cls As String
    cli As String
    anl As String
    sec As String
    desc As String
    sol As String
    reg_dt As Date
    cls_dt As Date
    has_cls As Boolean
    iso_year As Long
    iso_week As Long
    dif_Hour As Double
    topic_cluster As String
    tone_deviation_score As Long
    is_emotional As Boolean
    emotion_argument As String
    complexity As Long
    initial_grade As String
    final_grade As String
    grade_argument As String
    is_alternative As Boolean
    is_critical_incident As Boolean
End Type

Public Function GetSystemInstruction() As String
    Dim p As String
    p = "Вы — ведущий системный аналитик 1С. Проведите аудит пакета обращений." & vbCrLf & _
        "1. topic_cluster: конкретная суть сбоя/операции (без абстракций)." & vbCrLf & _
        "2. final_grade: ПЛОХО, ХОРОШО, ОТЛИЧНО." & vbCrLf & _
        "3. tone_deviation_score: от 0 (деловой) до 3 (оскорбления)." & vbCrLf & _
        "4. complexity: 1-5. is_alternative: true/false. is_critical_incident: true/false."
    GetSystemInstruction = p
End Function