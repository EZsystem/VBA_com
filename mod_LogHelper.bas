Attribute VB_Name = "mod_LogHelper"
'�܂��A�b�v���Ă��Ȃ���B�@2025/05/07�@13:43
'-------------------------------------
' Module: mod_LogHelper
' �����@�FAccess���ʃ��O�o�̓��W���[���ifrm_LogConsole �����j
' �쐬���F2025/05/07
' �X�V���F-
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' �֐��� : AppendLog
' ����   : frm_LogConsole �Ƀ��b�Z�[�W��ǋL���ĕ\��
'=================================================
Public Sub AppendLog(ByVal msg As String)
    On Error Resume Next

    Dim f As Form
    Set f = Forms("frm_LogConsole")

    ' --- �t�H�[�����J���Ă��Ȃ���ΊJ�� ---
    If f Is Nothing Then
        DoCmd.OpenForm "frm_LogConsole", acNormal
        Set f = Forms("frm_LogConsole")
    End If

    ' --- �ǋL���� ---
    Dim curLog As String
    curLog = Nz(f!txtLogOutput.value, "")
    f!txtLogOutput.value = curLog & Format(Now, "yyyy/mm/dd hh:nn:ss") & " > " & msg & vbCrLf

    ' --- �X�N���[���𖖔��ֈړ� ---
    f!txtLogOutput.SelStart = Len(f!txtLogOutput.Text)
End Sub

'=================================================
' �֐��� : ClearLog
' ����   : ���O�o�̓t�H�[���̃��O�����ׂď���
'=================================================
Public Sub ClearLog()
    On Error Resume Next
    Forms("frm_LogConsole")!txtLogOutput.value = ""
End Sub

