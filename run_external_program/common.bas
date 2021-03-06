Attribute VB_Name = "common"
Option Explicit

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
        (ByVal lpMutexAttributes As Long, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String) As Long
Declare Function WaitForSingleObject Lib "KERNEL32.DLL" _
        (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long
Declare Function ReleaseMutex Lib "kernel32" _
        (ByVal hMutex As Long) As Long
Declare Function GetLastError Lib "KERNEL32.DLL" () As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
        ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
        
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'--- Win32 API 定数の宣言 ---
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const INFINITE As Long = &HFFFFFFFF
Public Const WAIT_FAILED As Long = &HFFFFFFFF

Private Const CNFSHTNAM As String = "設定"

Private s_cnf_sht_da() As String

Enum MutexId
    en_pomodoro
    en_eventlog
    en_mutexid_cnt
End Enum

Private s_mutex_for_pomodoro(en_mutexid_cnt) As Long
Private Const s_mutex_namelist As String = "pomodoro,eventlog"

Public Function init_trclog(level As en_log_lvl, Optional cf As Boolean = True) As Boolean
    If set_logger_prm("logger_sheet.initialize_log_", "logger_sheet.output_log_", _
            "logger_sheet.finalize_log_", IIf(cf = True, "clear", ""), level) = True Then
        init_trclog = initialize_log
    End If
End Function

Public Function output_trace_log(level As en_log_lvl, message As String) As Boolean
    output_trace_log = output_log(level, message)
End Function

Public Sub final_trclog()
    finalize_log
End Sub

Private Sub tes_is_integer_string()
'    MsgBox is_integer_string("123")
'    MsgBox is_integer_string("-5")
    MsgBox is_integer_string("1.2")
    MsgBox is_integer_string("abc")
End Sub

Public Function is_integer_string(v As String) As Boolean
    On Error GoTo Err
    If v = CStr(CInt(v)) Then
        is_integer_string = True
    End If
Err:
End Function

Public Function create_mutex(mutex As MutexId, Optional force As Boolean = False) As Boolean
    On Error GoTo Err
    Application.EnableEvents = False    'イベント抑止
                                        '(スレッドには効果がないかも)
    Dim errlogmsg As String
    Do
        If s_mutex_for_pomodoro(mutex) <> 0 Then
            If force = True Then
                create_mutex = True
            End If
            Exit Do
        End If
        s_mutex_for_pomodoro(mutex) = CreateMutex(0, 0, Split(s_mutex_namelist, ",")(mutex))
        If s_mutex_for_pomodoro(mutex) <> 0 Then
            create_mutex = True
        Else
            errlogmsg = "CreateMutex(): errcod=" & CStr(GetLastError)
        End If
    Loop While False
Err:
    If Err.Number <> 0 Then
        errlogmsg = Err.Description
    End If
    Application.EnableEvents = True
    If Len(errlogmsg) > 0 Then
        Debug.Print "ERROR>> create_mutex(): " & errlogmsg
    End If
End Function

Public Function close_mutex(mutex As MutexId) As Boolean
    On Error GoTo Err
    Application.EnableEvents = False
    Dim errlogmsg As String
    Do
        If s_mutex_for_pomodoro(mutex) = 0 Then
            'nop
        ElseIf CloseHandle(s_mutex_for_pomodoro(mutex)) <> 0 Then
            s_mutex_for_pomodoro(mutex) = 0
        Else
            errlogmsg = "CloseHandle(): errcod=" & CStr(GetLastError)
            Exit Do
        End If
        close_mutex = True
    Loop While False
Err:
    If Err.Number <> 0 Then
        errlogmsg = Err.Description
    End If
    Application.EnableEvents = True
    If Len(errlogmsg) > 0 Then
        Debug.Print "ERROR>> close_mutex(): " & errlogmsg
    End If
End Function

Public Function lock_mutex(mutex As MutexId) As Boolean
    On Error GoTo Err
    Dim ret As Long
    ret = WaitForSingleObject(s_mutex_for_pomodoro(mutex), INFINITE)
    If ret = 0 Then
        lock_mutex = True
    Else
        Debug.Print "ERROR>> lock_mutex(): WaitForSingleObject(): errcod=" & CStr(GetLastError)
    End If
Err:
    If Err.Number <> 0 Then
        Debug.Print "ERROR>> lock_mutex(): " & Err.Description
    End If
End Function

Public Function unlock_mutex(mutex As MutexId) As Boolean
    On Error GoTo Err
'    mutex_unlock = True: Exit Function   'forデバッグ
    If ReleaseMutex(s_mutex_for_pomodoro(mutex)) <> 0 Then
        unlock_mutex = True
    End If
Err:
    If Err.Number <> 0 Then
        Debug.Print "ERROR>> unlock_mutex(): " & Err.Description
    End If
End Function

Public Function get_bottom_row(so As Worksheet, tr As Long, c As Long) As Long
    get_bottom_row = -1
    On Error GoTo Err
    If Len(so.Cells(tr, c).value) = 0 Or _
            Len(so.Cells(tr + 1, c).value) = 0 Then
        get_bottom_row = tr
    Else
        Range(so.Cells(tr, c), so.Cells(tr, c)).Select
        Selection.End(xlDown).Select
        get_bottom_row = Selection.Row
    End If
Err:
    If Err.Number <> 0 Then
        Debug.Print "get_bottom_row(): " & Err.Description
    End If
End Function

Public Function get_sheet_data( _
        so As Worksheet, _
        tr As Long, _
        br As Long, _
        lc As Long, _
        rc As Long, _
        da() As String) As Boolean
    On Error GoTo Err
    Dim a As Variant
    a = so.Range(so.Cells(tr, lc), so.Cells(br, rc))
    Dim i As Long, j As Long, mi As Long, mj As Long
    mi = br - tr
    mj = rc - lc
    ReDim da(0 To mi, 0 To mj)
    For i = 0 To mi
        For j = 0 To mj
            da(i, j) = a(i + 1, j + 1)
        Next
    Next
    get_sheet_data = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "get_sheet_data(): " & Err.Description
    End If
End Function

Private Sub tes_load_config()
    MsgBox load_config
End Sub

Public Function load_config() As Boolean
    On Error GoTo Err
    Dim so As Worksheet: Set so = ThisWorkbook.Sheets(CNFSHTNAM)
    so.Activate
    Dim tr As Long, br As Long, lc As Long, rc As Long: tr = 2: lc = 1: rc = 2
    br = get_bottom_row(so, tr, 1)
    If br <> -1 Then
        If get_sheet_data(so, tr, br, lc, rc, s_cnf_sht_da) = True Then
            load_config = True
        End If
    End If
Err:
    If Err.Number <> 0 Then
        Debug.Print "load_config(): " & Err.Description
    End If
End Function

Private Sub tes_get_confvalue()
    Dim ret As Boolean, v As String
    If load_config = True Then
        ret = get_confvalue("pomodoro", v)
        MsgBox "ret=" & CStr(ret) & ", v=" & v
        ret = get_confvalue("pomodoro_alerm_file：", v)
        MsgBox "ret=" & CStr(ret) & ", v=" & v
    End If
End Sub

Public Function get_confvalue(k As String, v As String) As Boolean
    On Error GoTo Err
    Dim i As Long, mi As Long
    mi = UBound(s_cnf_sht_da, 1)
    Dim pos As Long
    For i = 0 To mi
        pos = InStr(s_cnf_sht_da(i, 0), "：")
        If IIf(pos <> 0, Left(s_cnf_sht_da(i, 0), pos - 1), s_cnf_sht_da(i, 0)) = k Then
            v = s_cnf_sht_da(i, 1)
            get_confvalue = True
            Exit Function
        End If
    Next
    v = ""
Err:
    If Err.Number <> 0 Then
        Debug.Print "get_confvalue(): " & Err.Description
    End If
End Function

Function play_sound_file(sf As String) As Boolean
    On Error GoTo Err
    If mciSendString("Play " & sf, "", 0, 0) = 0 Then
        play_sound_file = True
    End If
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "play_sound_file(): " & Err.Description
    End If
End Function

Public Function load_text_file( _
        pth As String, _
        linary() As String, _
        linary_siz As Long, _
        Optional line_feed_code As String = vbLf) As Boolean
    On Error GoTo Err
'    Dim of As Boolean
'    Dim fn As Long
'    fn = FreeFile()
'    Open pth For Input As fn
'    of = True
'    Erase linary
'    Dim ls As String
'    linary_siz = 0
'    Do While Not EOF(fn)
'        Line Input #fn, ls
'        ReDim Preserve linary(0 To linary_siz)
'        linary(linary_siz) = ls
'        linary_siz = linary_siz + 1
'    Loop
    Dim buf As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(pth).OpenAsTextStream
            buf = .ReadAll
            .Close
        End With
    End With
    linary = Split(buf, line_feed_code)
    linary_siz = UBound(linary, 1) + 1
    load_text_file = True
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "load_text_file(): " & Err.Description
    End If
'    If of = True Then
'        Close fn
'    End If
End Function

Public Function load_csv_file( _
        pth As String, _
        datary() As String, _
        Optional limit As Long = -1) As Boolean
    On Error GoTo Err
    Dim linary() As String, linary_siz As Long
    If load_text_file(pth, linary, linary_siz) = False Then
        Exit Function
    ElseIf linary_siz <= 0 Then
        Exit Function
    End If
    Dim lindat() As String
    Dim i As Long, j As Long, mi As Long, mj As Long
    mi = linary_siz - 1
    ReDim datary(0 To mi, 0 To limit - 1)
    lindat = Split(linary(0), ",", limit)
    mj = UBound(lindat, 1)
    For j = 0 To mj
        datary(0, j) = lindat(j)
    Next
    For i = 1 To mi
        lindat = Split(linary(i), ",", limit)
        mj = UBound(lindat, 1)
        For j = 0 To mj
            datary(i, j) = lindat(j)
        Next
    Next
    load_csv_file = True
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "load_csv_file(): " & Err.Description
    End If
End Function

Public Function load_result_file( _
        rslpth As String, _
        cod As Long, _
        msg As String) As Boolean
    On Error GoTo Err
    Dim da() As String
    If load_csv_file(rslpth, da, 2) = False Then
        Exit Function
    End If
    cod = CLng(da(0, 0))
    msg = da(0, 1)
    load_result_file = True
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "load_result_file(): " & Err.Description
    End If
End Function

Public Function run_external_program( _
        cmdpth As String, _
        rslpth As String, _
        cod As Long, _
        msg As String, _
        Optional Timeout As Long = -1) As Boolean
    On Error GoTo Err
    Dim pid As Long
    pid = Shell(cmdpth, vbHide)
    Dim ph As Long
    ph = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    If ph <> 0 Then
        Dim ret As Long
        ret = WaitForSingleObject(ph, IIf(Timeout <> -1, Timeout, INFINITE))
        CloseHandle ph
        If ret <> 0 Then
            Exit Function
        End If
    End If
    If load_result_file(rslpth, cod, msg) = False Then
        Exit Function
    ElseIf cod = 0 Then
        run_external_program = True
    End If
Err:
    If Err.Number <> 0 Then
        output_trace_log errlog, "run_external_program(): " & Err.Description
    End If
End Function

