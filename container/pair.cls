VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_name As String
Private s_value As String

Property Get name() As String
    name = s_name
End Property

Property Get value() As String
    value = s_value
End Property

Public Sub set_(name As String, value As String)
    s_name = name
    s_value = value
End Sub
