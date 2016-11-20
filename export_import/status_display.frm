VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} status_display 
   Caption         =   "UserForm1"
   ClientHeight    =   3756
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   5490
   OleObjectBlob   =   "status_display.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "status_display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private s_fixed_status As String
Private s_line_count As Long

Public Sub show_start(ByVal title As String)
    s_fixed_status = ""
    Me.Caption = title
    Me.Label1.Caption = ""
    s_line_count = 0
    Me.Show
End Sub

Public Sub update_status(ByVal status As String, ByVal fixed As Boolean)
    If s_fixed_status <> "" Then
        Me.Label1.Caption = s_fixed_status & vbCrLf & status
    Else
        Me.Label1.Caption = status
    End If
    If fixed = True Then
        s_fixed_status = Me.Label1.Caption
        s_line_count = s_line_count + 1
        If s_line_count > 13 Then
            Dim p As Long
            p = InStr(s_fixed_status, vbCrLf)
            s_fixed_status = Right(s_fixed_status, Len(s_fixed_status) - p - (Len(vbCrLf) - 1))
        End If
    End If
    DoEvents
End Sub
