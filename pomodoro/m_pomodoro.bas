Attribute VB_Name = "m_pomodoro"
Option Explicit

Private Enum PomodoroStatus
    en_end = 0
    en_start
    en_stop
End Enum

Private s_status As PomodoroStatus

Private s_pomodoro As Long
Private s_short_braek As Long
Private s_long_break As Long
Private s_set_time As Double
Private s_set_timer As Long
Private s_pomodoro_alerm_file As String
Private s_time_label_obj As Object

Private s_stop_time As Double

Private Sub tes005()
    MsgBox "en_end=" & CStr(en_end) & ", en_start=" & CStr(en_start) & ", en_stop=" & CStr(en_stop)
End Sub

Private Sub tes_start_pomodoro()
    Dim ret As Boolean
    ret = start_pomodoro(-1, -1)
End Sub

Private Function set_config_data(pv As Long, sv As Long) As Boolean
    If load_config = False Then
        Exit Function
    End If
    Dim v As String
    If pv > 0 Then
        s_pomodoro = pv
    Else
        If get_confvalue("pomodoro", v) = False Then
            Exit Function
        End If
        s_pomodoro = CLng(v)
    End If
    If sv > 0 Then
        s_short_braek = sv
    Else
        If get_confvalue("short break", v) = False Then
            Exit Function
        End If
        s_short_braek = CLng(v)
    End If
    If get_confvalue("long break", v) = False Then
        Exit Function
    End If
    s_long_break = CLng(v)
    If get_confvalue("pomodoro_alerm_file", s_pomodoro_alerm_file) = False Then
        Exit Function
    End If
    s_set_timer = 0
    set_config_data = True
End Function

Private Sub tes_timer_string()
    MsgBox timer_string(5)
    MsgBox timer_string(81)
End Sub

Private Function timer_string(m As Long) As String
    timer_string = Right("00" & CStr(CLng(m / 60)), 2) & ":" & Right("00" & CStr(m Mod 60), 2) & ":00"
End Function

Private Sub tes002()
    Dim d1 As Date, d2 As Date
    d1 = TimeValue("00:01:00")
    d2 = TimeValue("00:01:23")
    Dim d3 As Date: d3 = CDate(d2 - d1)
    MsgBox Right("0" & CStr(Hour(d3)), 2) & ":" & Right("0" & CStr(Minute(d3)), 2) & ":" & _
            Right("0" & CStr(Second(d3)), 2)
End Sub

Private Sub set_timer(Optional af As Boolean = True)
    On Error GoTo Err
    Dim lf As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Sub
    End If
    lf = True
    Do
        If s_status <> en_start Then
            Exit Do
        End If
        If af = True Then
            play_sound_file s_pomodoro_alerm_file
        End If
        If s_set_timer = s_pomodoro And s_short_braek <> 0 Then
            s_set_timer = s_short_braek
        Else
            s_set_timer = s_pomodoro
        End If
        output_trace_log inflog, "set_timer(): NEXTタイマー=" & CStr(s_set_timer) & "分"
        s_set_time = Now + TimeValue(timer_string(s_set_timer))
        Application.OnTime s_set_time, "set_timer"
    Loop While False
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "set_timer(): " & Err.Description
    End If
    If lf = True Then
        unlock_mutex en_pomodoro
    End If
End Sub

Private Function clear_timer() As Boolean
    On Error GoTo Err
    If s_status = en_start Then
        Application.OnTime s_set_time, "set_timer", , False
    Else
        Exit Function
    End If
    clear_timer = True
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "clear_timer(): " & Err.Description
    End If
End Function

Public Function start_pomodoro(pv As Long, sv As Long, Optional sa As Boolean = True) As Boolean
    On Error GoTo Err
    If create_mutex(en_pomodoro, True) = False Then
        Exit Function
    End If
    Dim lf As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    lf = True
    Do
        If s_status = en_start Then
            clear_timer
            s_status = en_end   '一旦、終了状態にする
        End If
        If set_config_data(pv, sv) = False Then
            Exit Do
        End If
        s_status = en_start
        set_timer sa
        start_pomodoro = True
    Loop While False
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "start_pomodoro(): " & Err.Description
    End If
    If lf = True Then
        unlock_mutex en_pomodoro
    End If
End Function

Public Function end_pomodoro() As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    If clear_timer = True Then
        'nop
    End If
    s_set_timer = 0
    s_status = en_end
    end_pomodoro = True
    unlock_mutex en_pomodoro
End Function

'Private Sub tes001_settimer()
'    play_sound_file "E:\etc\sound\Alarm01.wav"
'    Dim st As Double
'    st = Now + TimeValue("00:00:05")
'    Application.OnTime st, "tes001_settimer"
'End Sub
'
'Private Sub tes001_1()
'    tes001_settimer
'    Debug.Print "hello_001_start\n"
'    Application.EnableEvents = False
'    Dim i As Long, j As Long
'    For i = 0 To 100000
'    For j = 0 To 200000
'    Next
'    Next
'    Debug.Print "hello_001_end\n"
'    Application.EnableEvents = True
'End Sub

Private Sub tes001_2()
    Debug.Print "hello_002\n"
    Application.EnableEvents = True
End Sub

Public Function is_pomodoro_start() As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    is_pomodoro_start = IIf(s_status = en_start, True, False)
    unlock_mutex en_pomodoro
End Function

Public Function is_pomodoro_end() As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    is_pomodoro_end = IIf(s_status = en_end, True, False)
    unlock_mutex en_pomodoro
End Function

Public Function is_pomodoro_stop() As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    is_pomodoro_stop = IIf(s_status = en_stop, True, False)
    unlock_mutex en_pomodoro
End Function

'Private Sub tes_get_remain_pomodoro_time()
'    MsgBox CDate(CDate("00:25:00") - CDate("00:01:10"))
'    MsgBox get_remain_pomodoro_time
'End Sub

Public Function get_remain_pomodoro_time() As String
    Dim d As Double
    d = CDate(s_set_time) - Now
    get_remain_pomodoro_time = Right("0" & CStr(Hour(d)), 2) & ":" & Right("0" & CStr(Minute(d)), 2) & ":" & _
            Right("0" & CStr(Second(d)), 2)
End Function

Public Sub start_pomodoro_disp(lo As Object)
    If lock_mutex(en_pomodoro) = False Then
        Exit Sub
    End If
    Set s_time_label_obj = lo
    unlock_mutex en_pomodoro
    update_remain_time
End Sub

Private Sub update_remain_time()
    If lock_mutex(en_pomodoro) = False Then
        Exit Sub
    End If
    If Not s_time_label_obj Is Nothing Then
        Debug.Print "update_remain_time() call!!"
        s_time_label_obj.Caption = get_remain_pomodoro_time
        Dim t As Double: t = Now + TimeValue("00:00:01")
        Application.OnTime t, "update_remain_time"
    End If
    unlock_mutex en_pomodoro
End Sub

Public Sub end_pomodoro_disp()
    Dim lf As Boolean
    If lock_mutex(en_pomodoro) = True Then
        lf = True
    End If
    end_pomodoro_disp_
    If lf = True Then
        unlock_mutex en_pomodoro
    End If
End Sub

Private Sub end_pomodoro_disp_()
    Set s_time_label_obj = Nothing
End Sub

'停止処理
Public Function stop_pomodoro() As Boolean
    On Error GoTo Err
    Dim lf As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    lf = True
    Do
        If s_status <> en_start Then
            Exit Do
        End If
        clear_timer
        end_pomodoro_disp_
        s_stop_time = TimeValue(get_remain_pomodoro_time)
        s_status = en_stop
        stop_pomodoro = True
    Loop While False
Err:
    If Err.Number <> 0 Then
        Debug.Print "stop_pomodoro(): " & Err.Description
    End If
    If lf = True Then
        unlock_mutex en_pomodoro
    End If
End Function

'再開処理
Public Function resume_pomodoro() As Boolean
    On Error GoTo Err
    Dim lf As Boolean
    If lock_mutex(en_pomodoro) = False Then
        Exit Function
    End If
    lf = True
    Do
        If s_status <> en_stop Then
            Exit Do
        End If
        output_trace_log inflog, "resume_pomodoro(): NEXTタイマー=" & CStr(s_set_timer) & "分"
        s_set_time = Now + s_stop_time
        Application.OnTime s_set_time, "set_timer"
        s_status = en_start
        resume_pomodoro = True
    Loop While False
Err:
    If Err.Number <> 0 Then
        Debug.Print "resume_pomodoro(): " & Err.Description
    End If
    If lf = True Then
        unlock_mutex en_pomodoro
    End If
End Function

