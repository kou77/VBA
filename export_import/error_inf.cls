VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "error_inf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_number As Integer
Private s_source As String
Private s_description As String

'Public Sub Class_Initialize()
'    'nop
'End Sub
'
'Public Sub Class_Terminate()
'    'nop
'End Sub

Public Sub set_errinf(number As Integer, source As String, description As String)
    s_number = number
    s_source = source
    s_description = description
End Sub

Public Sub get_errinf(number As Integer, source As String, description As String)
    number = s_number
    source = s_source
    description = s_description
End Sub

