Attribute VB_Name = "logger_debug"
Option Explicit

Public Function initialize_log_(prm As String) As Boolean
    On Error GoTo Err
    'ç°ÇÃÇ∆Ç±ÇÎ(Ç®ÇªÇÁÇ≠Ç∏Ç¡Ç∆)nop
    initialize_log_ = True
Err:
    If Err.number <> 0 Then
        Debug.Print "initialize_log_(): " & Err.description
    End If
End Function

Public Function output_log_(time As String, level As String, message As String) As Boolean
    On Error GoTo Err
    Debug.Print time & "," & level & "," & message
    output_log_ = True
Err:
    If Err.number <> 0 Then
        Debug.Print "output_log_(): " & Err.description
    End If
End Function

Public Function finalize_log_() As Boolean
    On Error GoTo Err
    'ç°ÇÃÇ∆Ç±ÇÎ(Ç®ÇªÇÁÇ≠Ç∏Ç¡Ç∆)nop
    finalize_log_ = True
Err:
    If Err.number <> 0 Then
        Debug.Print "finalize_log_(): " & Err.description
    End If
End Function

