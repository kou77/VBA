Attribute VB_Name = "logger"
Option Explicit

Public Enum en_log_lvl
    errlog = -1
    wnglog = -2
    inflog = 0
    'debug: 1, 2, 3...
End Enum

Private s_i_fnc_str As String
Private s_o_fnc_str As String
Private s_f_fnc_str As String
Private s_i_prm_str As String
Private s_level As Long

Private s_errlog_count As Long
Private s_wnglog_count As Long
Private s_inflog_count As Long
Private s_first_err_message As String
Private s_first_err_time As String

Public Function set_logger_prm( _
        i_fnc_str As String, _
        o_fnc_str As String, _
        f_fnc_str As String, _
        i_prm_str As String, _
        level As Long) As Boolean
    s_i_fnc_str = i_fnc_str
    s_o_fnc_str = o_fnc_str
    s_f_fnc_str = f_fnc_str
    s_i_prm_str = i_prm_str
    s_level = IIf(level > 0, level, 0)
    set_logger_prm = True
End Function

Private Sub tes001()
    MsgBox "[" & s_i_fnc_str & "]"
    MsgBox Len(s_i_fnc_str)
End Sub

Private Sub tes002()
    On Error GoTo Err
    Dim i As Integer: i = 10 / 0
Err:
    MsgBox Err.description
End Sub

Public Function initialize_log(Optional prm As String = "", Optional callback As String = "") As Boolean
    On Error GoTo Err
    Dim cb As String: cb = IIf(Len(callback) <> 0, callback, s_i_fnc_str)
    If Len(cb) = 0 Then
        Exit Function
    End If
    Dim p As String: p = IIf(prm = "none", "", IIf(Len(prm) <> 0, prm, s_i_prm_str))
    If Application.Run(cb, p) = True Then
        s_errlog_count = 0
        s_wnglog_count = 0
        s_inflog_count = 0
        s_first_err_message = ""
        s_first_err_time = ""
        initialize_log = True
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "initialize_log(): " & Err.description
    End If
End Function

Public Function output_log( _
        level As en_log_lvl, _
        message As String, _
        Optional from As String = "", Optional callback As String = "") As Boolean
    On Error GoTo Err
    Dim cb As String: cb = IIf(Len(callback) <> 0, callback, s_o_fnc_str)
    If Len(cb) = 0 Then
        Exit Function
    End If
    If level > s_level Then
        output_log = True
        Exit Function
    End If
    Dim ts As String
    ts = format(Date, "yyyy/mm/dd ") & _
            format(time, "hh:mm:ss.") & Trim(CStr(Fix((CDbl(Timer) - Fix(CDbl(Timer))) * 1000)))
    Dim ls As String: ls = get_level_string(level)
    Dim ms As String: ms = IIf(Len(from) > 0, from & "(): ", "") & message
    If Application.Run(cb, ts, ls, ms) = True Then
        If level = errlog Then
            s_errlog_count = s_errlog_count + 1
            If s_first_err_message = "" Then
                s_first_err_message = ms
                s_first_err_time = ts
            End If
        ElseIf level = wnglog Then
            s_wnglog_count = s_wnglog_count + 1
        ElseIf level = inflog Then
            s_inflog_count = s_inflog_count + 1
        End If
        output_log = True
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "output_log(): " & Err.description
    End If
End Function

Private Function get_level_string(level As en_log_lvl) As String
    If level = errlog Then
        get_level_string = "ERROR"
    ElseIf level = wnglog Then
        get_level_string = "WARNING"
    ElseIf level = inflog Then
        get_level_string = "INFOMATION"
    Else
        get_level_string = "DEBUG" & CStr(level)
    End If
End Function

Public Function finalize_log(Optional callback As String = "") As Boolean
    On Error GoTo Err
    Dim cb As String: cb = IIf(Len(callback) <> 0, callback, s_f_fnc_str)
    If Len(cb) = 0 Then
        Exit Function
    End If
    s_errlog_count = 0
    s_wnglog_count = 0
    s_inflog_count = 0
    s_first_err_message = ""
    s_first_err_time = ""
    finalize_log = Application.Run(cb)
Err:
    If Err.number <> 0 Then
        Debug.Print "finalize_log(): " & Err.description
    End If
End Function

Public Sub get_log_count(errcnt As Long, wngcnt As Long, infcnt As Long)
    errcnt = s_errlog_count
    wngcnt = s_wnglog_count
    infcnt = s_inflog_count
End Sub

Public Sub get_first_errlog(em As String, et As String)
    em = s_first_err_message
    et = s_first_err_time
End Sub

Public Property Get error_count() As Long
    error_count = s_errlog_count
End Property

