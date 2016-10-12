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

Public Sub add(e As String)
    Dim i As Long: i = maxidx + 1
    ReDim Preserve s_ary(0 To i)
    s_ary(i) = e
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

