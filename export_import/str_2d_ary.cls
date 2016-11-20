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

Public Sub add_s1ary_rcd(r As str_1d_ary)
    Dim i As Long: i = maxidx + 1
    ReDim Preserve s_ary(0 To i)
    Set s_ary(i) = r
End Sub

Public Sub add_rcd(r() As String)
    Dim r_ As str_1d_ary: Set r_ = New str_1d_ary
    r_.set_data r
    Me.add_s1ary_rcd r_
End Sub

Public Sub insert_rcd(ByVal i As Long, r() As String)
    Dim i_ As Long, mi As Long: mi = maxidx
    If i > mi Then raise ERR_SYSTEM, "str_2d_ary::insert_rcd() ", "不正インデックス"
    Dim r_ As str_1d_ary: Set r_ = New str_1d_ary
    r_.set_data r
    ReDim Preserve s_ary(0 To mi + 1)
    For i_ = mi To i Step -1
        s_ary(i_ + 1) = s_ary(i_)
    Next
    s_ary(i) = r_
End Sub

Public Sub delete_rcd(ByVal i As Long)
    Dim i_ As Long, mi As Long: mi = maxidx
    If i > mi Then raise ERR_SYSTEM, "str_2d_ary::delete_rcd() ", "不正インデックス"
    If mi = 0 Then
        Erase s_ary
        Exit Sub
    End If
    For i_ = i + 1 To mi
        s_ary(i_ - 1) = s_ary(i_)
    Next
    ReDim Preserve s_ary(0 To mi - 1)
End Sub

Public Function get_(i As Long) As str_1d_ary
    On Error GoTo Err
    Set get_ = s_ary(i)
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "str_2d_ary::get_(): ", Err.description
    End If
End Function

Public Sub clear()
    Erase s_ary
End Sub

Public Sub set_data(da() As String)
    Me.clear
'    Dim i As Long, mi As Long, j As Long, mj As Long
'    mi = UBound(da, 1): mj = UBound(da, 2)
'    Dim s1da As str_1d_ary
'    For i = 0 To mi
'        Set s1da = New str_1d_ary
'        For j = 0 To mj
'            s1da.add da(i, j)
'        Next
'        Me.add s1da
'    Next
    add_data da
End Sub

Public Sub add_data(da() As String)
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = UBound(da, 1): mj = UBound(da, 2)
    Dim s1da As str_1d_ary
    For i = 0 To mi
        Set s1da = New str_1d_ary
        For j = 0 To mj
            s1da.add da(i, j)
        Next
        Me.add_s1ary_rcd s1da
    Next
End Sub

Public Function search(ByVal kc As Long, ByVal v As String) As Long
    Dim i As Long, mi As Long: mi = maxidx
    For i = 0 To mi
        If s_ary(i).get_(kc) = v Then
            search = i
            Exit Function
        End If
    Next
    search = -1
End Function

Public Sub get_data(da() As String)
    Dim i As Long, mi As Long: mi = maxidx
    If mi = -1 Then
        Erase da
        Exit Sub
    End If
    Dim rcd() As String
    Dim j As Long, mj As Long
    For i = 0 To mi
        If s_ary(i).maxidx > mj Then mj = s_ary(i).maxidx
    Next
    ReDim da(0 To mi, 0 To mj)
    For i = 0 To mi
        mj = s_ary(i).maxidx
        For j = 0 To mj
            da(i, j) = s_ary(i).get_(j)
        Next
    Next
End Sub

Public Sub sort(ByVal ci As Long, Optional reverse As Boolean = False, Optional kind As String = "string")
    Dim i As Long, j As Long, mi As Long: mi = maxidx
    Dim r As str_1d_ary, p As Long
'    For i = 0 To mi
'        Set r = s_ary(i)
'        For j = i + 1 To mi
'            If kind = "date" Then
'                If IIf(reverse = False, CDate(s_ary(j).get_(ci)) < CDate(r.get_(ci)), CDate(s_ary(j).get_(ci)) > CDate(r.get_(ci))) Then _
'                    Set r = s_ary(j): p = j
'            ElseIf kind = "string" Then
'                If IIf(reverse = False, s_ary(j).get_(ci) < r.get_(ci), s_ary(j).get_(ci) > r.get_(ci)) Then _
'                    Set r = s_ary(j): p = j
'            Else    'kind="double"
'                If IIf(reverse = False, CDbl(s_ary(j).get_(ci)) < CDbl(r.get_(ci)), CDbl(s_ary(j).get_(ci)) > CDbl(r.get_(ci))) Then _
'                    Set r = s_ary(j): p = j
'            End If
'        Next
'        If i <> p Then
'            For j = p To i + 1 Step -1
'                Set s_ary(j) = s_ary(j - 1)
'            Next
'            Set s_ary(i) = r
'        End If
'    Next
'    Exit Sub
    Dim cmp As Long
    For i = 0 To mi
        Set r = s_ary(i)
        p = i
        For j = i + 1 To mi
            cmp = compare(s_ary(j).get_(ci), r.get_(ci), kind)
            If IIf(reverse = False, cmp < 0, cmp > 0) Then
                Set r = s_ary(j)
                p = j
            End If
        Next
        If i <> p Then
            For j = p To i + 1 Step -1
                Set s_ary(j) = s_ary(j - 1)
            Next
            Set s_ary(i) = r
        End If
    Next
End Sub

Public Sub set_record(ByVal idx As Long, rcd() As String)
    s_ary(idx).set_data rcd
End Sub

Public Sub reverse()
    Dim i As Long, mi As Long: mi = maxidx
    If mi = -1 Then Exit Sub
    Dim da() As str_1d_ary: ReDim da(0 To mi)
    For i = 0 To mi
        da(i) = s_ary(mi - i)
    Next
    s_ary = da
End Sub

