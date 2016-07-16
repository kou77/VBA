VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_pomodoro 
   Caption         =   "pomodoro"
   ClientHeight    =   3630
   ClientLeft      =   96
   ClientTop       =   384
   ClientWidth     =   4608
   OleObjectBlob   =   "f_pomodoro.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "f_pomodoro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'�u�J�n�v�{�^������
Private Sub CommandButton1_Click()
    On Error GoTo Err
    Application.ScreenUpdating = False
    Dim so As Worksheet: Set so = ActiveSheet
    init_trclog 0
    Dim pv As Long, sv As Long: pv = -1: sv = -1
    If is_integer_string(Me.TextBox1.value) = True Then
        pv = CLng(Me.TextBox1.value)
    End If
    If is_integer_string(Me.TextBox2.value) = True Then
        sv = CLng(Me.TextBox2.value)
    End If
    Dim af As Boolean: af = IIf(Me.CheckBox1.value, False, True)
    Dim ret As Boolean: ret = start_pomodoro(pv, sv, af)
    If ret = True Then
        Me.CommandButton4.Enabled = True
    End If
Err:
    post_processing "f_pomodoro::CommandButton1_Click", ret, Err.Number, Err.Description, so, False
End Sub

'�u�I���v�{�^������
Private Sub CommandButton2_Click()
    On Error GoTo Err
    Application.ScreenUpdating = False
    Dim so As Worksheet: Set so = ActiveSheet
    init_trclog 0
    Dim ret As Boolean: ret = end_pomodoro
    If ret = True Then
        Me.CommandButton4.Caption = "��~"
        Me.CommandButton4.Enabled = False
        Me.Label3.Visible = False
        m_pomodoro.end_pomodoro_disp
    End If
Err:
    post_processing "f_pomodoro::CommandButton2_Click", ret, Err.Number, Err.Description, so, False
End Sub

'�u����v�{�^������
Private Sub CommandButton3_Click()
    Hide_
End Sub

'�u��~/�ĊJ�v�{�^������
Private Sub CommandButton4_Click()
    If is_pomodoro_end = True Then
        Exit Sub    '�������ɂ͗��Ȃ��͂�
    End If
    On Error GoTo Err
    Application.ScreenUpdating = False
    Dim so As Worksheet: Set so = ActiveSheet
    init_trclog 0
    Dim ret As Boolean
    If is_pomodoro_start = True Then       '�u��~�v�{�^�������H
        ret = stop_pomodoro
        If ret = True Then
            Me.CommandButton4.Caption = "�ĊJ"
        End If
    Else                                   '�u�ĊJ�v�{�^������
        ret = resume_pomodoro
        If ret = True Then
            Me.CommandButton4.Caption = "��~"
        End If
    End If
Err:
    post_processing "f_pomodoro::CommandButton4_Click", ret, Err.Number, Err.Description, so, False
    Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Activate()
    'mutex�����쐬�̏ꍇ�ɍ쐬����
    If create_mutex(en_pomodoro, True) = False Then
        Exit Sub
    End If
    If is_pomodoro_start = True Then
        m_pomodoro.start_pomodoro_disp Me.Label3
        Me.Label3.Visible = True
        Me.CommandButton4.Caption = "��~"
        Me.CommandButton4.Enabled = True
    ElseIf is_pomodoro_end = True Then
        Me.Label3.Visible = False
        Me.CommandButton4.Caption = "��~"
        Me.CommandButton4.Enabled = False
    Else
        'nop
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "�m����n�{�^�����g�p���Ă�������"
        Cancel = True
    End If
End Sub

Private Sub post_processing(method As String, ret As Boolean, en As Long, ed As String, so As Worksheet, _
        Optional hf As Boolean = True)
    Dim okf As Boolean
    If en <> 0 Then
        output_trace_log errlog, method & "():" & ed
    ElseIf ret = True Then
        output_trace_log inflog, method & "(): ����I��"
        okf = True
    Else
        output_trace_log errlog, method & "(): �G���[�I��"
    End If
    final_trclog
    If Not so Is Nothing Then
        so.Activate
    End If
    Application.ScreenUpdating = True
    If okf = False Then
        'nop
    ElseIf hf = True Then
        Hide_
    ElseIf is_pomodoro_start = True Then
        m_pomodoro.start_pomodoro_disp Me.Label3
        Me.Label3.Visible = True
    End If
End Sub

Private Sub Hide_()
    m_pomodoro.end_pomodoro_disp
    Me.Hide
End Sub
