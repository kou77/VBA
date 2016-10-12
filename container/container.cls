VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_datary As str_2d_ary

Public Sub Class_Initialize()
    Set s_datary = New str_2d_ary
End Sub

Public Sub clear()
    s_datary.clear
End Sub

Public Sub set_data(da() As String)
    If s_datary Is Nothing Then
        Set s_datary = New str_2d_ary
    Else
        s_datary.clear
    End If
    Dim rcd As str_1d_ary
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = UBound(da, 1): mj = UBound(da, 2)
    For i = 0 To mi
        Set rcd = New str_1d_ary
        For j = 0 To mj
            rcd.add da(i, j)
        Next
        s_datary.add rcd
    Next
End Sub

Public Sub get_data(da() As String)
    Dim i As Long, mi As Long, j As Long, mj As Long
    mi = s_datary.maxidx
    mj = -1
    For i = 0 To mi
        If s_datary.get_(i).maxidx > mj Then mj = s_datary.get_(i).maxidx
    Next
    ReDim da(0 To mi, 0 To mj)
    For i = 0 To mi
        mj = s_datary.get_(i).maxidx
        For j = 0 To mj
            da(i, j) = s_datary.get_(i).get_(j)
        Next
    Next
End Sub

Public Sub add_rcd(da() As String)
    Dim rcd As str_1d_ary: Set rcd = New str_1d_ary
    Dim i As Long, mi As Long: mi = UBound(da, 1)
    For i = 0 To mi
        rcd.add da(i)
    Next
    s_datary.add rcd
End Sub

Public Function get_column(ByVal ri As Long, ByVal ci As Long) As String
    On Error GoTo Err
    get_column = s_datary.get_(ri).get_(ci)
Err:
    If Err.number <> 0 Then
        raise ERR_EXCEPTION, "container::get_column(): ", Err.description
    End If
End Function

Public Sub get_record(ByVal ri As Long, rcd() As String)
    Dim i As Long, mi As Long: mi = s_datary.get_(ri).maxidx
    ReDim rcd(0 To mi)
    For i = 0 To mi
        rcd(i) = s_datary.get_(ri).get_(i)
    Next
End Sub

Public Function search_record(ByVal ci As Long, ByVal v As String, Optional ByVal si As Long = 0) As Long
    On Error GoTo Err
    search_record = -1
    Dim i As Long, mi As Long: mi = maxidx
    For i = si To mi
        If s_datary.get_(i).get_(ci) = v Then
            search_record = i
            Exit Function
        End If
    Next
Err:
End Function

Property Get maxidx() As Long
    On Error GoTo Err
    maxidx = -1
    maxidx = s_datary.maxidx
Err:
End Property


