VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "str_2d_ary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_ary() As str_1d_ary

Property Get maxidx() As Long
    On Error GoTo Err
    maxidx = -1
    maxidx = UBound(s_ary, 1)
Err:
End Property

Public Sub add(e As str_1d_ary)
    Dim i As Long: i = maxidx + 1
    ReDim Preserve s_ary(0 To i)
    Set s_ary(i) = e
End Sub

Public Function get_(i As Long) As str_1d_ary
    On Error GoTo Err
    Set get_ = s_ary(i)
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "str_1d_ary::get_(): ", Err.description
    End If
End Function

Public Sub clear()
    Erase s_ary
End Sub

Public Sub set_data(da() As String)
    Me.clear
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = UBound(da, 1): mj = UBound(da, 2)
    Dim s1da As str_1d_ary
    For i = 0 To mi
        Set s1da = New str_1d_ary
        For j = 0 To mj
            s1da.add da(i, j)
        Next
        Me.add s1da
    Next
End Sub

