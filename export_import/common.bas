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

Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" ( _
        ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'--- Win32 API 定数の宣言 ---
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const INFINITE As Long = &HFFFFFFFF
Public Const WAIT_FAILED As Long = &HFFFFFFFF

Public Const CNFSHTNAM As String = "設定"

Public Const OPERATION_SHEET_TOP_ROW As Long = 9
Public Const OPERATION_SHEET_KEY_CLM As Long = 4

Public Const QMARK_LIST As String = """'!@*&|"

Public Const ERR_SYSTEM As Long = 1
Public Const ERR_EXCEPTION As Long = 2

'Private s_cnf_sht_da() As String
Private s_cnf_sht_da As container

Private s_errinf As error_inf
Private s_source_inf As String

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
    Set s_errinf = New error_inf
End Function

Public Function output_trace_log(level As en_log_lvl, message As String) As Boolean
    output_trace_log = output_log(level, message)
End Function

Public Sub final_trclog()
    finalize_log
End Sub

Private Sub tes_is_integer_string()
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
    If read_csv_file(rslpth, da, 2) = False Then
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

Public Function is_symbol_str(s As String) As Boolean
    Dim i As Long, mi As Long, c As String
    mi = Len(s)
    If mi = 0 Then
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
        Exit Function   '空文字で復帰
    ElseIf s = "<" Then
        get_r_qmark = ">"
    ElseIf s = "[" Then
        get_r_qmark = "]"
    ElseIf s = "{" Then
        get_r_qmark = "}"
    ElseIf s = "(" Then
        get_r_qmark = ")"
'    ElseIf s = """" Or s = "'" Then
    ElseIf InStr(QMARK_LIST, s) <> 0 Then
        get_r_qmark = s
    Else
        get_r_qmark = ""
    End If
End Function

'引数wf: 警告レベルのログ出力がある場合も、カウント文字列を組み立てる
Public Function error_cnt_str(Optional wf As Boolean = False) As String
    Dim errcnt As Long, wngcnt As Long, infcnt As Long
    get_log_count errcnt, wngcnt, infcnt
    If errcnt = 0 And wngcnt = 0 Then
    ElseIf errcnt > 0 Or wf = True Then
        error_cnt_str = ": errcnt=" & CStr(errcnt) & ", wngcnt=" & CStr(wngcnt) & ", infcnt=" & CStr(infcnt)
    End If
End Function

Public Function write_container2file(c As container) As Boolean
    Dim da() As String
    c.get_data da
    Dim pth As String: pth = Application.ThisWorkbook.path & "\anchor.txt"
    write_container2file = write_csv_file(pth, da, qmflg:=True)
End Function

Private Sub tes_display_result_message()
    On Error GoTo Err
    Dim a As Long: a = 1 / 0
Err:
    display_result_message errobj:=Err
End Sub

Public Sub display_result_message(Optional ByVal ret As Boolean = True, Optional ByVal errobj As ErrObject = Nothing, _
        Optional ByVal error_only As Boolean = True, Optional ByVal en As Long = 0, Optional ByVal ed As String = "", _
        Optional ByVal ecs As String = "")
    Dim lc As String: lc = IIf(ecs <> "", ecs, error_cnt_str)
    Dim ms As String
    If ret = False Or error_count > 0 Then
        ms = "エラー終了"
    ElseIf en <> 0 Then
        ms = "エラー終了" & IIf(ed <> "", "(" & ed & ")", "")
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
    If lc <> "" Then ms = ms & lc
    MsgBox ms
End Sub

Public Sub set_pair_ary(ByVal v As String, pa As pair_ary)
    v = Trim(v)
    Dim pn As String, pv As String
    Do While Len(v) > 0
        extract_pair v, pn, pv
        If v <> "" Then
            If Left(v, 1) <> "," Then raise ERR_SYSTEM, "set_pair_ary(): " & get_source_inf, _
                    "不正な区切り文字[" & Left(v, 1) & "]"
            v = Trim(Right(v, Len(v) - 1))
        End If
        If pn <> "" Then
            If is_symbol_str(pn) = False Then
                raise ERR_SYSTEM, "set_pair_ary(): " & get_source_inf, "無効なシンボル[" & pn & "]"
            End If
            pa.add pn, pv
        End If
    Loop
End Sub

Private Sub extract_pair(v As String, pn As String, pv As String)
    On Error GoTo Err
    pn = "": pv = ""
    If Left(v, 1) = "," Then Exit Sub
    Dim p1 As Long, p2 As Long: p1 = InStr(v, "="): p2 = InStr(v, ",")
    If p1 = 0 And p2 = 0 Then   '１つの変数名
        pn = v
        v = ""
        Exit Sub
    ElseIf p2 <> 0 And (p1 > p2 Or p1 = 0) Then     '値なしの変数名、継続あり
        pn = Trim(Left(v, p2 - 1))
        v = Trim(Right(v, Len(v) - p2 + 1))
        Exit Sub
    End If
    'ここから引数('='(イコール))ありの場合
    pn = Trim(Left(v, p1 - 1))
    v = Trim(Right(v, Len(v) - p1))
    If v = "" Then Exit Sub
    If Left(v, 1) = "," Then Exit Sub   '値が空
    Dim r_qm As String: r_qm = get_r_qmark(Left(v, 1))
    If r_qm <> "" Then  '引用符を調べる
        p1 = InStr(2, v, r_qm)
    Else
        p1 = 0
    End If
    If p1 <> 0 Then     '引用符付き値あり
        pv = Mid(v, 2, p1 - 2)
        v = Trim(Right(v, Len(v) - p2))
    ElseIf p2 > 0 Then  '引用符なし値、継続あり
        pv = Trim(Left(v, p1 - 1))
        v = Trim(Right(v, Len(v) - p2 - 1))
    Else                '引用符なし値、継続なし
        pv = v
        v = ""
    End If
Err:
    If Err.number <> 0 Then raise ERR_EXCEPTION, "extract_pair(): ", Err.description
End Sub

Public Function extract_value(ByVal s As String, v As String, rp As Long, Optional ByVal eov As String = ",") As Boolean
    On Error GoTo Err
    Dim s_ As String: s_ = s
    Dim p As Long
    If s_ = "" Then
        v = ""
    Else
        Dim r_qm As String: r_qm = get_r_qmark(Left(s_, 1))
        If r_qm <> "" Then p = InStr(2, s_, r_qm)
        If p <> 0 Then
            v = Mid(s_, 2, p - 2)
            s_ = skip_space(Right(s_, Len(s_) - p))
        Else
            If Len(eov) > 0 Then
                p = InStr(s, eov)
                If p > 0 Then
                    v = Left(s_, p - 1)
                    s_ = Right(s_, Len(s_) - p + 1)
                Else
                    v = s_
                    s_ = ""
                End If
            Else
                v = s_
                s_ = ""
            End If
        End If
    End If
    If s_ <> "" Then
        If eov <> "" And Left(s_, Len(eov)) <> eov Then Exit Function
        rp = Len(s) - Len(s_) + Len(eov) + 1
    Else
        rp = Len(s) + 1
    End If
    extract_value = True
Err:
    If Err.number <> 0 Then output_log errlog, "extract_value(): " & Err.description
End Function

Public Function extract_variable_value2pair(ByVal s As String, p As pair, rp As Long, Optional ByVal eov As String = ",") As Boolean
    On Error GoTo Err
    Dim pos As Long: pos = InStr(s, "=")
    Dim n As String, v As String
    If pos = 0 Then
        If eov <> "" Then
            pos = InStr(s, eov)
        End If
        If pos <> 0 Then
            n = Left(s, pos - 1)
            rp = pos + Len(eov)
        Else
            n = s
            rp = Len(s) + 1
        End If
    Else
        n = Left(s, pos - 1)
        Dim s_ As String: s_ = Right(s, Len(s) - pos)
        If extract_value(s_, v, rp, eov) = False Then Exit Function
        rp = pos + rp
    End If
    If is_symbol_str(n) = False Then Exit Function
    p.set_ n, v
    extract_variable_value2pair = True
Err:
    If Err.number <> 0 Then output_log errlog, "extract_variable_value2pair(): " & Err.description
End Function

Public Function skip_space(ByVal s As String) As String
    Dim i As Long, mi As Long: mi = Len(s)
    Dim c As String
    For i = 1 To mi
        c = Mid(s, i, 1)
        If c <> " " And c <> vbTab And c <> vbLf And c <> vbCr Then
            skip_space = Right(s, Len(s) - i + 1)
            Exit Function
        End If
    Next
    skip_space = ""
End Function

Public Function extract_value_list(ByVal s As String, va As str_1d_ary, Optional ByVal separator As String = ",") As Boolean
    va.clear
    Dim s_ As String: s_ = s
    Dim rp As Long, v As String
    Do While True
        If extract_value(s_, v, rp, separator) = False Then Exit Function
        va.add v
        If rp > Len(s_) Then Exit Do
        s_ = skip_space(Right(s_, Len(s_) - rp - -1))
    Loop
    extract_value_list = True
End Function

Public Function get_qm_chr(ByVal v As String, qm_chr As String) As Boolean
    Dim i As Long, mi As Long: mi = Len(QMARK_LIST)
    For i = 1 To mi
        If InStr(v, Mid(QMARK_LIST, i, 1)) = 0 Then
            qm_chr = Mid(QMARK_LIST, i, 1)
            get_qm_chr = True
            Exit Function
        End If
    Next
End Function

'a_f1 + a_f2 => a_t
Public Function join_str2dary(a_f1 As str_2d_ary, a_f2 As str_2d_ary, a_t As str_2d_ary) As Boolean
    On Error GoTo Err
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = a_f1.maxidx
    a_t.clear
    Dim rcd() As String
    For i = 0 To mi
        mj = a_f1.get_(i).maxidx
        ReDim rcd(0 To mj)
        For j = 0 To mj
            rcd(j) = a_f1.get_(i).get_(j)
        Next
        a_t.add_rcd rcd
    Next
    mi = a_f2.maxidx
    For i = 0 To mi
        mj = a_f2.get_(i).maxidx
        ReDim rcd(0 To mj)
        For j = 0 To mj
            rcd(j) = a_f2.get_(i).get_(j)
        Next
        a_t.add_rcd rcd
    Next
    join_str2dary = True
Err:
    If Err.number <> 0 Then output_log errlog, "join_str2dary(): " & Err.description
End Function

'c_f1 + c_f2 => c_t
Public Function join_container(c_f1 As container, c_f2 As container, c_t As container) As Boolean
    On Error GoTo Err
    c_t.set_container_data c_f1
    c_t.add_container_data c_f2
    join_container = True
Err:
    If Err.number <> 0 Then output_log errlog, "join_container(): " & Err.description
End Function

Public Function remove_space(ByVal s As String) As String
    Dim s_ As String
    s_ = Replace(s, vbCrLf, "")
    s_ = Replace(s_, vbCr, "")
    s_ = Replace(s_, vbLf, "")
    s_ = Replace(s_, " ", "")
    s_ = Replace(s_, vbTab, "")
    remove_space = s_
End Function

Public Function dirname(ByVal path As String, Optional m As String = "dir") As String
    Dim p As Long
    If m = "dir" Then
        p = InStrRev(path, "/")
    Else
        p = InStrRev(path, "\")
    End If
    If p > 0 Then
        dirname = Left(path, p - 1)
    Else
        dirname = ""
    End If
End Function

Public Function basename(ByVal path As String, Optional m As String = "dir") As String
    Dim p As Long
    If m = "dir" Then
        p = InStrRev(path, "/")
    Else
        p = InStrRev(path, "\")
    End If
    If p > 0 Then
        basename = Right(path, Len(path) - p)
    Else
        basename = path
    End If
End Function

Private Sub tes005()
    MsgBox get_current_date_and_time_string
    MsgBox get_current_date_and_time_string(True)
End Sub

Public Function get_current_date_and_time_string(Optional ByVal ms_grant As Boolean = False) As String
    get_current_date_and_time_string = format(Date, "yyyy/mm/dd ") & _
            format(time, "hh:mm:ss") & _
            IIf(ms_grant = True, "." & Trim(CStr(Fix((CDbl(Timer) - Fix(CDbl(Timer))) * 1000))), "")
End Function

Public Function get_file_date_last_modified(ByVal path As String, ds As String) As Boolean
    On Error GoTo Err
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ds = fso.GetFile(path).DateLastModified
    get_file_date_last_modified = True
Err:
    If Err.number <> 0 Then output_log errlog, "get_file_date_last_modified(): " & Err.description
    Set fso = Nothing
End Function

Public Function start_status(ByVal title As String) As Boolean
    On Error GoTo Err
    status_display.show_start (title)
    start_status = True
Err:
End Function

Public Function update_status(ByVal status As String, ByVal fixed As Boolean) As Boolean
    On Error GoTo Err
    status_display.update_status status, fixed
    update_status = True
Err:
End Function

Public Function end_status() As Boolean
    On Error GoTo Err
    status_display.Hide
    end_status = True
Err:
End Function

Public Function compare(ByVal v1 As String, ByVal v2 As String, Optional kind As String = "string") As Long
    If v1 = v2 Then compare = 0: Exit Function
    If kind = "date" Then
        compare = IIf(CDate(v1) > CDate(v2), 1, -1)
    ElseIf kind = "double" Then
        compare = IIf(CDbl(v1) > CDbl(v2), 1, -1)
    Else    'kind="double"
        If kind <> "string" Then
            output_log wnglog, "compare(): 不正なkindの指定(kind=" & kind & ")"
        End If
        compare = IIf(v1 > v2, 1, -1)
    End If
End Function


