; ウィンドウタイトル
WinTitle := "FAX送信ポップアップ"

; ウィンドウの表示を最大10秒待つ
WinWait, %WinTitle%,, 10
if ErrorLevel {
    MsgBox, エラー: "%WinTitle%" が見つかりませんでした。
    ExitApp
}

; ウィンドウをアクティブに
WinActivate, %WinTitle%
WinWaitActive, %WinTitle%

; コントロール名（必要に応じて修正）
FaxNumberControl := "Edit2"
SendButton       := "Button5"
CloseButton      := "Button19"

; 番号入力
ControlSetText, %FaxNumberControl%, 00362068221, %WinTitle%

Sleep, 10

; 送信
ControlClick, %SendButton%, %WinTitle%
Send {Space}

Sleep, 5

; 完了ボタン
ControlClick, %CloseButton%, %WinTitle%
Send {Space}

TrayTip,, FAX送信スクリプトが完了しました。, 3
ExitApp
