VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'「import」ボタン押下
Private Sub CommandButton1_Click()
    Dim ret As Boolean
    On Error GoTo Err
    ret = import(Cells(5, 3).Value)
Err:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    ElseIf ret = True Then
        MsgBox "正常終了"
    Else
        MsgBox "エラー終了"
    End If
End Sub

Private Function import(fld_pth As String) As Boolean
    If Dir(fld_pth, vbDirectory) = "" Then
        MsgBox "importフォルダ(" & fld_pth & ")が存在しない"
        Exit Function
    End If
    On Error GoTo Err
    Dim fil As String
    fil = Dir(fld_pth & "\*")
    Do While fil <> ""
        If is_import_file(fil) = True Then
            If import_1c(fld_pth & "\" & fil) = False Then
                Exit Function
            End If
        End If
        fil = Dir
    Loop
    import = True
Err:
    If Err.Number <> 0 Then
        MsgBox "import(): " & Err.Description
    End If
End Function

'「export」ボタン押下
Private Sub CommandButton2_Click()
    Dim ret As Boolean
    On Error GoTo Err
    ret = export(Cells(6, 3).Value)
Err:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    ElseIf ret = True Then
        MsgBox "正常終了"
    Else
        MsgBox "エラー終了"
    End If
End Sub

Private Function export(fld_pth As String) As Boolean
    On Error GoTo Err
    If Dir(fld_pth, vbDirectory) = "" Then
        MsgBox "export(" & fld_pth & ")が存在しない"
        Exit Function
    End If
    Dim ret As Boolean
    Dim c As VBComponent
    For Each c In ThisWorkbook.VBProject.VBComponents
        If is_export_component(c) = True Then
            c.export fld_pth & "\" & c.Name & "." & get_cextstr(c.Type)
        End If
    Next
    export = True
Err:
    If Err.Number <> 0 Then
        MsgBox "export(): " & Err.Description
    End If
End Function

'※対象外コンポーネントの場合は空文字を返す
Private Function get_cextstr(ext As vbext_ComponentType) As String
    If ext = vbext_ct_ClassModule Then
        get_cextstr = "cls"
    ElseIf ext = vbext_ct_StdModule Then
        get_cextstr = "bas"
    End If
End Function

Private Function import_1c(pth As String) As Boolean
    On Error GoTo Err
    Dim fil As String: fil = basename(pth)
    Dim c As VBComponent
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Name & "." & get_cextstr(c.Type) = fil Then
            Exit For
        End If
    Next
    If Not c Is Nothing Then    'コンポーネントが存在？
                                '(存在したら一旦削除)
        ThisWorkbook.VBProject.VBComponents.Remove c
    End If
    ThisWorkbook.VBProject.VBComponents.import pth
    import_1c = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "import_1c(): " & Err.Description
    End If
End Function

Private Function is_import_file(fil As String) As Boolean
    If (Right(fil, Len(".cls")) = ".cls" And Left(fil, Len("Sheet")) <> "Sheet") Or _
            Right(fil, Len(".bas")) = ".bas" Then
        is_import_file = True
    End If
End Function

Private Function is_export_component(c As VBComponent) As Boolean
    If c.Type = vbext_ct_ClassModule Or _
            c.Type = vbext_ct_StdModule Then
        is_export_component = True
    End If
End Function

Private Function basename(pth As String) As String
    Dim pos As Long: pos = InStrRev(pth, "\")
    basename = Right(pth, Len(pth) - pos)
End Function
