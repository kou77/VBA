Attribute VB_Name = "logger_file"
Option Explicit

Private Const DEFAULT_FILE_NAME As String = "log.txt"

Private s_file_name As String
Private s_fd As Long
Private s_opnflg As Boolean

Public Function initialize_log_(prm As String) As Boolean
    On Error GoTo Err
    Dim v As String
    If check_parameter(prm, "file_name", v) = True Then
        s_file_name = v
    Else
        s_file_name = DEFAULT_FILE_NAME
    End If
    s_fd = FreeFile
    Dim pth As String: pth = Application.Workbooks(Application.ActiveWorkbook.Name).Path & "\" & s_file_name
    Open pth For Append As s_fd
    s_opnflg = True
    initialize_log_ = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "initialize_log_(): " & Err.Description
    End If
End Function

Public Function output_log_(time As String, level As String, message As String) As Boolean
    On Error GoTo Err
    If s_opnflg = False Then
        Exit Function
    End If
    Print #s_fd, time & "," & level & "," & message
    output_log_ = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "output_log_(): " & Err.Description
    End If
End Function

Public Function finalize_log_() As Boolean
    On Error GoTo Err
    If s_opnflg = False Then
        Exit Function
    End If
    Close s_fd
    finalize_log_ = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "finalize_log_(): " & Err.Description
    End If
End Function


