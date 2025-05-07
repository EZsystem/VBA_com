Attribute VB_Name = "mod_LogHelper"
'まだアップしていないよ。　2025/05/07　13:43
'-------------------------------------
' Module: mod_LogHelper
' 説明　：Access共通ログ出力モジュール（frm_LogConsole 向け）
' 作成日：2025/05/07
' 更新日：-
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' 関数名 : AppendLog
' 説明   : frm_LogConsole にメッセージを追記して表示
'=================================================
Public Sub AppendLog(ByVal msg As String)
    On Error Resume Next

    Dim f As Form
    Set f = Forms("frm_LogConsole")

    ' --- フォームが開いていなければ開く ---
    If f Is Nothing Then
        DoCmd.OpenForm "frm_LogConsole", acNormal
        Set f = Forms("frm_LogConsole")
    End If

    ' --- 追記処理 ---
    Dim curLog As String
    curLog = Nz(f!txtLogOutput.value, "")
    f!txtLogOutput.value = curLog & Format(Now, "yyyy/mm/dd hh:nn:ss") & " > " & msg & vbCrLf

    ' --- スクロールを末尾へ移動 ---
    f!txtLogOutput.SelStart = Len(f!txtLogOutput.Text)
End Sub

'=================================================
' 関数名 : ClearLog
' 説明   : ログ出力フォームのログをすべて消去
'=================================================
Public Sub ClearLog()
    On Error Resume Next
    Forms("frm_LogConsole")!txtLogOutput.value = ""
End Sub

