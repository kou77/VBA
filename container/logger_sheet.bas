Attribute VB_Name = "logger_sheet"
Option Explicit

Private Const DEFAULT_SHEET_NAME As String = "é¿çsÉçÉO"

Private s_sheet_name As String
Private s_output_row As Long

Private Sub tes_initialize_log_()
    Dim ret As Boolean
    ret = initialize_log_("sheet_name=test")
End Sub

Public Function initialize_log_(prm As String) As Boolean
    On Error GoTo Err
    Dim so1 As Worksheet, so2 As Worksheet
    Set so1 = ActiveSheet
    Dim v As String
    If check_parameter(prm, "sheet_name", v) = True Then
        s_sheet_name = v
    Else
        s_sheet_name = DEFAULT_SHEET_NAME
    End If
    Dim so As Variant, ff As Boolean
    For Each so In ThisWorkbook.Sheets
        If so.name = s_sheet_name Then
            ff = True
            Exit For
        End If
    Next
    Do
'        If Not so Is Empty Then
        If ff = True Then
            Set so2 = ThisWorkbook.Sheets(s_sheet_name)
            so2.Activate
            If check_parameter(prm, "clear", v) = True Then
                If clear_logsht = False Then
                    Exit Function
                End If
                s_output_row = 1
            Else
                s_output_row = get_next_logrow(so2)
                If s_output_row = -1 Then
                    Exit Do
                End If
            End If
        Else
            Set so2 = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            so2.name = s_sheet_name
            so2.Activate
            Cells.Select
            Selection.NumberFormatLocal = "@"
            so2.Cells(1, 1).Select
            s_output_row = 1
        End If
        initialize_log_ = True
    Loop While False
Err:
    If Err.number <> 0 Then
        Debug.Print "initialize_log_(): " & Err.description
    End If
    If Not so1 Is Nothing Then
        so1.Activate
    End If
End Function

Public Function output_log_(time As String, level As String, message As String) As Boolean
    On Error GoTo Err
    ThisWorkbook.Sheets(s_sheet_name).Cells(s_output_row, 1) = time
    ThisWorkbook.Sheets(s_sheet_name).Cells(s_output_row, 2) = level
    ThisWorkbook.Sheets(s_sheet_name).Cells(s_output_row, 3) = message
    s_output_row = s_output_row + 1
    output_log_ = True
Err:
    If Err.number <> 0 Then
        Debug.Print "output_log_(): " & Err.description
    End If
End Function

Public Function finalize_log_() As Boolean
    On Error GoTo Err
    'ç°ÇÃÇ∆Ç±ÇÎnop
    finalize_log_ = True
Err:
    If Err.number <> 0 Then
        Debug.Print "finalize_log_(): " & Err.description
    End If
End Function

Private Sub tes001()
'    MsgBox Sheets("test")
'    Dim so As Worksheet: Set so = Sheets("test")
    Dim s As Variant
    For Each s In ThisWorkbook.Sheets
        MsgBox s.name
    Next
End Sub

Private Sub tes_clear_logsht()
    s_sheet_name = "test"
    clear_logsht
End Sub

'êÊì™ÇÃÉJÉâÉÄÇæÇØÇí≤Ç◊ÇÈ
Private Function get_next_logrow(so As Worksheet) As Long
    On Error GoTo Err
    get_next_logrow = -1
    If Len(so.Cells(1, 1).value) = 0 Then
        get_next_logrow = 1
    ElseIf Len(so.Cells(2, 1).value) = 0 Then
        get_next_logrow = 2
    Else
        Range(so.Cells(1, 1), so.Cells(1, 1)).Select
        Selection.End(xlDown).Select
        get_next_logrow = Selection.Row + 1
    End If
Err:
    If Err.number <> 0 Then
        Debug.Print "clear_logsht(): " & Err.description
    End If
End Function

Private Function clear_logsht() As Boolean
    On Error GoTo Err
    Dim so As Worksheet: Set so = ThisWorkbook.Sheets(s_sheet_name)
    If Len(so.Cells(1, 1).value) <> 0 Then
        Dim r As Long: r = get_next_logrow(so) - 1
        If r > 0 Then
            Rows("1:" & CStr(r)).Select
            Selection.Delete Shift:=xlUp
        End If
        so.Cells(1, 1).Select
    End If
    clear_logsht = True
Err:
    If Err.number <> 0 Then
        Debug.Print "clear_logsht(): " & Err.description
    End If
End Function

