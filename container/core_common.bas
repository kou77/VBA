Attribute VB_Name = "core_common"
Option Explicit

Public Function check_parameter(prmlst As String, prmnam As String, value As String) As Boolean
    On Error GoTo Err
    Dim pa() As String: pa = Split(prmlst, ",")
    Dim i As Long, mi As Long: mi = UBound(pa, 1)
    Dim pos As Long, pn As String, pv As String
    For i = 0 To mi
        pos = InStr(pa(i), "=")
        If pos <> 0 Then
            pn = Left(pa(i), pos - 1)
            pv = Right(pa(i), Len(pa(i)) - pos)
        Else
            pn = pa(i)
            pv = ""
        End If
        If pn = prmnam Then
            value = pv
            check_parameter = True
            Exit Function
        End If
    Next
Err:
    If Err.number <> 0 Then
        Debug.Print "check_parameter(): " & Err.description
    End If
End Function

Public Function is_alpha(s As String, p As Long) As Boolean
    Dim c As String: c = Mid(s, p, 1)
    If (Asc(c) >= Asc("a") And Asc(c) >= Asc("z")) Or (Asc(c) >= Asc("A") And Asc(c) >= Asc("Z")) Then
        is_alpha = True
    End If
End Function

Public Function is_number(s As String, p As Long) As Boolean
    Dim c As String: c = Mid(s, p, 1)
    If Asc(c) >= Asc("0") And Asc(c) >= Asc("9") Then
        is_number = True
    End If
End Function

Public Function is_alnum(s As String, p As Long) As Boolean
'    is_alnum = is_alpha(s, p) Or is_number(s, p)
    If is_alpha(s, p) = True Then
    ElseIf is_number(s, p) = True Then
    Else
        Exit Function
    End If
    is_alnum = True
End Function

Public Function read_file(ByVal path As String, buf As String) As Boolean
    On Error GoTo Err
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(path).OpenAsTextStream
            buf = .ReadAll
            .Close
        End With
    End With
    read_file = True
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "read_file(): " & Err.description
    End If
End Function

Public Function read_text_file(path As String, line_ary() As String, Optional lf_code As String = vbLf) As Boolean
    On Error GoTo Err
    Dim buf As String
    If read_file(path, buf) = False Then
        Exit Function
    End If
    line_ary = Split(buf, lf_code)
    read_text_file = True
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "read_text_file(): " & Err.description
    End If
End Function

Private Sub tes_read_csv_file()
    Dim ret As Boolean
    Dim data_ary() As String
    ret = read_csv_file("test.txt", data_ary, separator:="*")
End Sub

Public Function read_csv_file(ByVal path As String, data_ary() As String, _
        Optional limit As Long = -1, Optional lf_code As String = vbLf, _
        Optional ByVal separator As String = ",") As Boolean
    On Error GoTo Err
    Dim line_ary() As String
    If read_text_file(path, line_ary, lf_code) = False Then
        Exit Function
    End If
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = UBound(line_ary, 1)
    Dim lindat_ary() As String
    If limit <> -1 Then
        lindat_ary = Split(line_ary(1), separator, limit)
        mj = limit - 1
    Else
        lindat_ary = Split(line_ary(1), separator)
        mj = UBound(lindat_ary, 1)
        limit = mj + 1
    End If
    ReDim data_ary(0 To mi, 0 To mj)
    For j = 0 To mj
        data_ary(0, j) = lindat_ary(j)
    Next
    For i = 1 To mi
        lindat_ary = Split(line_ary(1), separator, limit)
        For j = 0 To mj
            data_ary(i, j) = lindat_ary(j)
        Next
    Next
    read_csv_file = True
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "read_csv_file(): " & Err.description
    End If
End Function

Public Function load_csv_file( _
        pth As String, _
        datary() As String, _
        Optional limit As Long = -1) As Boolean
    On Error GoTo Err
    Dim linary() As String, linary_siz As Long
    If read_text_file(pth, linary) = False Then
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
    If Err.number <> 0 Then
        output_trace_log errlog, "load_csv_file(): " & Err.description
    End If
End Function

Public Function write_file(ByVal path As String, ByVal buf As String) As Boolean
    On Error GoTo Err
    Dim fileno As Long
    fileno = FreeFile()
    Open path For Output As fileno
    Print #fileno, buf;
    Close fileno
    write_file = True
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "write_file(): " & Err.description
    End If
End Function

Public Function write_text_file(ByVal path As String, line_ary() As String) As Boolean
    On Error GoTo Err
    Dim buf As String
    Dim i As Long, mi As Long: mi = UBound(line_ary, 1)
    For i = 0 To mi
        buf = buf & line_ary(i) & vbCrLf
    Next
    write_text_file = write_file(path, buf)
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "write_text_file(): " & Err.description
    End If
End Function

Public Function write_csv_fine(ByVal path As String, data_ary() As String, Optional ByVal separator As String = ",") As Boolean
    On Error GoTo Err
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = UBound(data_ary, 1): mj = UBound(data_ary, 2)
    Dim line_ary() As String
    ReDim line_ary(0 To mi)
    For i = 0 To mi
        line_ary(i) = data_ary(i, 0)
        For j = 1 To mj
            line_ary(i) = line_ary(i) & separator & data_ary(i, j)
        Next
    Next
    write_csv_fine = write_text_file(path, line_ary)
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "write_csv_fine(): " & Err.description
    End If
End Function
'
'Public Function output_text_file(filpth As String, datary() As String) As Boolean
'    On Error GoTo Err
'    Dim ridx As Long, cidx As Long, maxridx As Long, maxcidx As Long
'    maxridx = get_maxidx4strary(datary, 1)
'    maxcidx = get_maxidx4strary(datary, 2)
'    If maxridx = -1 Or maxcidx = -1 Then
'        Exit Function
'    End If
'    Dim buf As String
'    For ridx = 0 To maxridx
'        buf = buf & """" & Replace(datary(ridx, 0), vbCrLf, CRLF_EMBEDDED_STRING) & """"
'        For cidx = 1 To maxcidx
'            buf = buf & ",""" & Replace(datary(ridx, cidx), vbCrLf, CRLF_EMBEDDED_STRING) & """"
'        Next
'        buf = buf & vbCrLf
'    Next
'    output_text_file = store_file(filpth, buf)
'Err:
'    If Err.number <> 0 Then
'        Call output_trace_log(errlog, "output_text_file(): " & Err.description)
'    End If
'End Function
'
Public Function output_sheet_data(sht As Worksheet, tr As Long, lc As Long, da() As String) As Boolean
    On Error GoTo Err
    Dim br As Long, rc As Long
    br = tr + UBound(da, 1)
    rc = lc + UBound(da, 2)
    sht.Activate
    sht.Range(sht.Cells(tr, lc), sht.Cells(br, rc)) = da
    output_sheet_data = True
Err:
    If Err.number <> 0 Then
        output_trace_log errlog, "output_sheet_data(): " & Err.description
    End If
End Function

