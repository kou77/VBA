Attribute VB_Name = "message"
Option Explicit

'�ȉ����b�Z�[�W�̓T���v��
Public Enum message_no
    '�V�X�e��
    en_system001 = 1
    en_system002
    en_system003
    en_system004
    en_system005
    '�ݒ�(�R���t�B�O)
    en_cnfinf001
    '�L���t�@�C��
    en_artfil001
    en_artfil002
    en_artfil003
    en_artfil004
    '�^�O�t�@�C��
    en_tagfil001
    'changelog
    en_chglog001
    en_chglog002
End Enum

Private Const SYSTEM_MESSAGE_LIST As String = _
    "�A���J�[�t�@�C���ǂݍ��ݎ��s" & vbCrLf & _
    "�A���J�[�t�@�C���s�����R�[�h(%%%p1%%%)" & vbCrLf & _
    "�J�e�S�����(cno=%%%p1%%%)��������Ȃ�" & vbCrLf & _
    "�L�����(ano=%%%p1%%%)��������Ȃ�" & vbCrLf & _
    "�Y���^�O(%%%p1%%%)�����݂��Ȃ�" & vbCrLf

Private Const CNFINF_MESSAGE_LIST As String = _
    "�ݒ���(%%%p1%%%)�擾�G���[" & vbCrLf

Private Const ARTFIL_MESSAGE_LIST As String = _
    "�Ή������Ă��Ȃ��I��pair�^�O(%%%p1%%%): current_tag=%%%p2%%%" & vbCrLf & _
    "�I���^�O(%%%p1%%%)�����݂��Ȃ�" & vbCrLf & _
    "������location(%%%p1%%%): current_tag=%%%p2%%%" & vbCrLf & _
    "���o�����x�����X�L�b�v(�����x��=%%%p1%%%, �w�背�x��=%%%p2%%%, ���o����=%%%p3%%%)" & vbCrLf

Private Const TAGFIL_MESSAGE_LIST As String = _
    "�Y���^�O(%%%p1%%%)�����݂��Ȃ�" & vbCrLf

Private Const CHGLOG_MESSAGE_LIST As String = _
    "��񂪕s�����Ă���" & vbCrLf & _
    "���I����񂠂�"

Private Const MESSAGE_LIST As String = _
    SYSTEM_MESSAGE_LIST & _
    CNFINF_MESSAGE_LIST & _
    ARTFIL_MESSAGE_LIST & _
    TAGFIL_MESSAGE_LIST & _
    CHGLOG_MESSAGE_LIST

Private s_message_ary As str_1d_ary

Private Const AA As String = "abc" & vbCrLf
Private Const BB As String = "777" & vbCrLf
Private Const CC As String = AA & BB

Private Sub tes001()
    MsgBox CStr(en_system002)
    MsgBox get_message(en_system001)
End Sub

Private Sub tes002()
    MsgBox CC
End Sub

Public Function get_message(ByVal mn As message_no, _
        Optional ByVal p1 As String = "", Optional ByVal p2 As String = "", Optional ByVal p3 As String = "", _
        Optional ByVal p4 As String = "", Optional ByVal p5 As String = "") As String
    If s_message_ary Is Nothing Then
        Dim da() As String: da = Split(MESSAGE_LIST, vbCrLf)
        Set s_message_ary = New str_1d_ary
        s_message_ary.set_data da
    End If
    Dim m As String, pc As Long
    If mn > s_message_ary.maxidx + 1 Then _
        get_message = "### ILIAGAL MESSAGE NO.[" & CStr(mn) & "] ###": Exit Function
    m = s_message_ary.get_(mn - 1)
    If InStr(m, "%%%p1%%%") = 0 Then
    ElseIf InStr(m, "%%%p2%%%") = 0 Then
        pc = 1
    ElseIf InStr(m, "%%%p3%%%") = 0 Then
        pc = 2
    ElseIf InStr(m, "%%%p4%%%") = 0 Then
        pc = 3
    ElseIf InStr(m, "%%%p5%%%") = 0 Then
        pc = 4
    Else
        pc = 5
    End If
    Do
        If pc = 0 Then Exit Do
        m = Replace(m, "%%%p1%%%", p1)
        If pc = 1 Then Exit Do
        m = Replace(m, "%%%p2%%%", p2)
        If pc = 2 Then Exit Do
        m = Replace(m, "%%%p3%%%", p3)
        If pc = 3 Then Exit Do
        m = Replace(m, "%%%p4%%%", p4)
        If pc = 4 Then Exit Do
        m = Replace(m, "%%%p5%%%", p5)
    Loop While False
    get_message = m
End Function

