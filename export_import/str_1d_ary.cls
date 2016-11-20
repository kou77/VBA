VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "str_1d_ary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_ary() As String

Property Get maxidx() As Long
    On Error GoTo Err
    maxidx = -1
    maxidx = UBound(s_ary, 1)
Err:
End Property

Public Sub add(ByVal v As String)
    Dim i As Long: i = maxidx + 1
    ReDim Preserve s_ary(0 To i)
    s_ary(i) = v
End Sub

Public Sub insert(ByVal i As Long, ByVal v As String)
    Dim i_ As Long, mi As Long: mi = maxidx
    If i > mi Then raise ERR_SYSTEM, "str_1d_ary::insert() ", "不正インデックス"
    ReDim Preserve s_ary(0 To mi + 1)
    For i_ = mi To i Step -1
        s_ary(i_ + 1) = s_ary(i_)
    Next
    s_ary(i) = v
End Sub

Public Sub delete(ByVal i As Long)
    Dim i_ As Long, mi As Long: mi = maxidx
    If i > mi Then raise ERR_SYSTEM, "str_1d_ary::delete() ", "不正インデックス"
    If mi = 0 Then
        Erase s_ary
        Exit Sub
    End If
    For i_ = i + 1 To mi
        s_ary(i_ - 1) = s_ary(i_)
    Next
    ReDim Preserve s_ary(0 To mi - 1)
End Sub

Public Function get_(i As Long) As String
    On Error GoTo Err
    get_ = s_ary(i)
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "str_1d_ary::get_(): ", Err.description
    End If
End Function

Public Sub clear()
    Erase s_ary
End Sub

Public Sub get_data(da() As String)
    Dim i As Long, mi As Long: mi = maxidx
    ReDim da(0 To mi)
    For i = 0 To mi
        da(i) = s_ary(i)
    Next
End Sub

Public Sub set_data(da() As String)
    On Error GoTo Err
    Dim i As Long, mi As Long: mi = UBound(da, 1)
    ReDim s_ary(0 To mi)
    For i = 0 To mi
        s_ary(i) = da(i)
    Next
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "str_1d_ary::set_data(): ", Err.description
    End If
End Sub

Public Function search(ByVal v As String) As Long
    Dim i As Long, mi As Long: mi = maxidx
    For i = 0 To mi
        If s_ary(i) = v Then
            search = i
            Exit Function
        End If
    Next
    search = -1
End Function

Public Sub reverse()
    Dim i As Long, mi As Long: mi = maxidx
    If mi = -1 Then Exit Sub
    Dim da() As String: ReDim da(0 To mi)
    For i = 0 To mi
        da(i) = s_ary(mi - i)
    Next
    s_ary = da
End Sub
