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

Public Const CNFSHTNAM As String = "設定"

Public Const OPERATION_SHEET_TOP_ROW As Long = 10
Public Const OPERATION_SHEET_KEY_CLM As Long = 3

'Private s_cnf_sht_da() As String
Private s_cnf_sht_da As container

Private s_errinf As error_inf
Private s_source_inf As String

Enum MutexId
    en_pomodoro
    en_eventlog
    en_mutexid_cnt
End Enum

Public Const ERR_NOTATION_DEFINF_001 As Long = 550
Public Const ERR_CONFIG As Long = 600
Public Const ERR_CATEGORY_INF As Long = 610
Public Const ERR_ARTICLE_INF As Long = 611
Public Const ERR_STRUCTURE_INF As Long = 612
Public Const ERR_INVALID_VARIABLE_VALUE As Long = 700
Public Const ERR_SYSTEM As Long = 900
Public Const ERR_EXCEPTION As Long = 999

Private s_mutex_for_pomodoro(en_mutexid_cnt) As Long
Private Const s_mutex_namelist As String = "pomodoro,eventlog"

Public Function init_trclog(level As en_log_lvl, Optional cf As Boolean = True) As Boolean
    If set_logger_prm("logger_sheet.initialize_log_", "logger_sheet.output_log_", _
            "logger_sheet.finalize_log_", IIf(cf = True, "clear", ""), level) = True Then
        init_trclog = initialize_log
    End If
    Set s_errinf = New error_inf
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
    If Err.number <> 0 Then
        errlogmsg = Err.description
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
    If Err.number <> 0 Then
        errlogmsg = Err.description
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
    If Err.number <> 0 Then
        Debug.Print "ERROR>> lock_mutex(): " & Err.description
    End If
End Function

Public Function unlock_mutex(mutex As MutexId) As Boolean
    On Error GoTo Err
'    mutex_unlock = True: Exit Function   'forデバッグ
    If ReleaseMutex(s_mutex_for_pomodoro(mutex)) <> 0 Then
        unlock_mutex = True
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "ERROR>> unlock_mutex(): " & Err.description
    End If
End Function

Public Function get_bottom_row(so As Worksheet, tr As Long, c As Long) As Long
    get_bottom_row = -1
    On Error GoTo Err
    so.Activate
    If Len(so.Cells(tr, c).value) = 0 Or _
            Len(so.Cells(tr + 1, c).value) = 0 Then
        get_bottom_row = tr
    Else
        Range(so.Cells(tr, c), so.Cells(tr, c)).Select
        Selection.End(xlDown).Select
        get_bottom_row = Selection.Row
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "get_bottom_row(): " & Err.description
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
    If Err.number <> 0 Then
        Debug.Print "get_sheet_data(): " & Err.description
    End If
End Function

Private Sub tes_load_config()
    MsgBox load_config
End Sub

Public Function load_config(Optional ByVal clear As Boolean = False) As Boolean
    On Error GoTo Err
    If clear = False And Not s_cnf_sht_da Is Nothing Then
        load_config = True
        Exit Function
    End If
    If s_cnf_sht_da Is Nothing Then
        Set s_cnf_sht_da = New container
    End If
    Dim so As Worksheet: Set so = ThisWorkbook.Sheets(CNFSHTNAM)
    so.Activate
    Dim tr As Long, br As Long, lc As Long, rc As Long: tr = 2: lc = 1: rc = 2
    br = get_bottom_row(so, tr, 1)
    If br <> -1 Then
        Dim da() As String
        If get_sheet_data(so, tr, br, lc, rc, da) = True Then
            s_cnf_sht_da.set_data da
            load_config = True
        End If
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "load_config(): " & Err.description
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

Public Function get_confvalue(ByVal k As String, v As String) As Boolean
    On Error GoTo Err
    Dim i As Long, mi As Long
    mi = s_cnf_sht_da.maxidx
    Dim pnam As String, pos As Long
    For i = 0 To mi
        pnam = s_cnf_sht_da.get_column(i, 0)
        pos = InStr(pnam, "：")
        If pos <> 0 Then pnam = Left(pnam, pos - 1)
        If pnam = k Then
            v = s_cnf_sht_da.get_column(i, 1)
            get_confvalue = True
            Exit Function
        End If
    Next
    v = ""
Err:
    If Err.number <> 0 Then
        Debug.Print "get_confvalue(): " & Err.description
    End If
End Function

Function play_sound_file(sf As String) As Boolean
    On Error GoTo Err
    If mciSendString("Play " & sf, "", 0, 0) = 0 Then
        play_sound_file = True
    End If
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "play_sound_file(): " & Err.description
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
    If Err.number <> 0 Then
        output_trace_log errlog, "load_result_file(): " & Err.description
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
    If Err.number <> 0 Then
        output_trace_log errlog, "run_external_program(): " & Err.description
    End If
End Function

Private Sub tes003()
    init_trclog inflog
    On Error GoTo Err
'    Dim v As Integer: v = 10 / 0
    tes003_1
Err:
    If Err.number <> 0 Then
        MsgBox error_checker(Err, "tes003")
    End If
    final_trclog
End Sub

Private Sub tes003_1()
    raise 550, "tes003_1(): ", "error111"
End Sub

Public Sub raise(number As Integer, source As String, discription As String)
    s_errinf.set_errinf number, source, discription
    Dim v As Integer: v = 10 / 0
End Sub

Public Function error_checker(Err As ErrObject, fn As String) As Boolean
    If Err.number = 0 Then
        error_checker = True
        Exit Function
    End If
    If Not s_errinf Is Nothing Then
        Dim number As Integer, source As String, discription As String
        s_errinf.get_errinf number, source, discription
        If number <> 0 Then
            output_log errlog, source & "error_code=" & CStr(number) & ": " & discription
            Exit Function
        End If
    End If
    output_log errlog, fn & "(): " & Err.description
End Function

Public Sub set_source_inf(source As String)
    s_source_inf = source
End Sub

Public Function get_source_inf() As String
    get_source_inf = s_source_inf
End Function

Private Sub tes002()
    Dim p As pair
    Set p = New pair
End Sub

Public Function is_symbol_str(s As String) As Boolean
    Dim i As Long, mi As Long, c As String
    mi = Len(s)
    If mi = -1 Then
        Exit Function
    End If
    c = Left(s, 1)
    If c <> "_" And is_alpha(c, 1) = False Then
        Exit Function
    End If
    For i = 2 To mi
        c = Mid(s, i, 1)
        If c <> "_" And is_alnum(c, 1) = False Then
            Exit Function
        End If
    Next
    is_symbol_str = True
End Function

Public Function get_r_qmark(s As String) As String
    If Len(s) <> 1 Then
        raise ERR_NOTATION_DEFINF_001, "notation_definf::Let kind(): " & get_source_inf, "記法定義:kind不正な設定値"    '★未実装
    ElseIf s = "<" Then
        get_r_qmark = ">"
    ElseIf s = "[" Then
        get_r_qmark = "]"
    ElseIf s = "{" Then
        get_r_qmark = "}"
    ElseIf s = "(" Then
        get_r_qmark = ")"
    Else
        get_r_qmark = s
    End If
End Function

Public Function error_cnt_str() As String
    Dim errcnt As Long, wngcnt As Long, infcnt As Long
    get_log_count errcnt, wngcnt, infcnt
    If errcnt > 0 Then
        error_cnt_str = ": errcnt=" & CStr(errcnt) & ", wngcnt=" & CStr(wngcnt) & ", infcnt=" & CStr(infcnt)
    End If
End Function

Public Function write_container2file(c As container) As Boolean
    Dim da() As String
    c.get_data da
    Dim pth As String: pth = Application.ThisWorkbook.path & "\anchor.txt"
    write_container2file = write_csv_fine(pth, da)
End Function

Private Sub tes_display_result_message()
    On Error GoTo Err
    Dim a As Long: a = 1 / 0
Err:
    display_result_message errobj:=Err
End Sub

Public Sub display_result_message(Optional ByVal ret As Boolean = True, Optional ByVal errobj As ErrObject = Nothing, _
        Optional ByVal error_only As Boolean = True)
    Dim ec As Long, wc As Long, ic As Long
    get_log_count ec, wc, ic
    Dim ms As String
    If ret = False Or ec > 0 Then
        ms = "エラー終了"
    ElseIf Not errobj Is Nothing Then
        If errobj.number <> 0 Then
            ms = "エラー終了"
        Else
            If error_only = True Then Exit Sub
            ms = "正常終了"
        End If
    Else
        ms = "正常終了"
    End If
    If ec > 0 Or wc > 0 Then ms = ms & "(errcnt=" & CStr(ec) & ", wngcnt=" & CStr(wc) & ", infcnt=" & CStr(ic)
    MsgBox ms
End Sub

