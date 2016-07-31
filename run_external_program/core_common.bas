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
    If Err.Number <> 0 Then
        Debug.Print "check_parameter(): " & Err.Description
    End If
End Function
