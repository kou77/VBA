VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lng_1d_ary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_ary() As Long

Property Get maxidx() As Long
    On Error GoTo Err
    maxidx = -1
    maxidx = UBound(s_ary, 1)
Err:
End Property

Public Sub add(e As Long)
    Dim i As Long: i = maxidx + 1
    ReDim Preserve s_ary(0 To i)
    s_ary(i) = e
End Sub

Public Function get_(i As Long) As Long
    On Error GoTo Err
    get_ = s_ary(i)
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "lng_1d_ary::get_(): ", Err.description
    End If
End Function

Public Sub clear()
    Erase s_ary
End Sub

Public Sub get_data(da() As Long)
    Dim i As Long, mi As Long: mi = maxidx
    ReDim da(0 To mi)
    For i = 0 To mi
        da(i) = s_ary(i)
    Next
End Sub

Public Sub set_data(da() As Long)
    On Error GoTo Err
    Dim i As Long, mi As Long: mi = UBound(da, 1)
    ReDim s_ary(0 To mi)
    For i = 0 To mi
        s_ary(i) = da(i)
    Next
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "lng_1d_ary::set_data(): ", Err.description
    End If
End Sub

Public Function search(ByVal v As Long) As Long
    Dim i As Long, mi As Long: mi = maxidx
    For i = 0 To mi
        If s_ary(i) = v Then
            search = i
            Exit Function
        End If
    Next
    search = -1
End Function

'Public Sub redim_(ByVal maxidx As Long)
'    'if maxidx < 0 then raiseš
'    ReDim Preserve s_ary(0 To maxidx)
'End Sub
