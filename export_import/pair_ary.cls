VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pair_ary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_pair_ary() As pair

Public Sub add(name As String, value As String)
    Dim idx As Long: idx = maxidx + 1
    ReDim Preserve s_pair_ary(0 To idx)
    Set s_pair_ary(idx) = New pair
    s_pair_ary(idx).set_ name, value
End Sub

Public Function maxidx() As Long
    On Error GoTo Err
    maxidx = -1
    maxidx = UBound(s_pair_ary, 1)
Err:
End Function

Public Function get_(idx As Long) As pair
    get_ = s_pair_ary(idx)
End Function

Public Function search(name As String, value As String) As Long
    Dim i As Long, mi As Long: mi = maxidx
    For i = 0 To mi
        If s_pair_ary(i).name = name Then
            value = s_pair_ary(i).value
            search = i
            Exit Function
        End If
    Next
    search = -1
End Function

