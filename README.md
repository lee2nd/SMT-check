# SMT-check

vba password : frank123

--------
2023/04/04 更新

1. 另存 txt
2. NG 第二次刷 QR code 才讓人輸入 PASS
3. 顯示全碼料號
4. now() 時間轉成純文字非公式
5. UI 文字改大一點
6. 若 OK 自動按 Enter 的功能

Private Sub TextBox1_Change()
        If allowAutoClick Then
            allowAutoClick = False ' 設定為 False，避免自動觸發事件
            CommandButton1_Click
            allowAutoClick = True ' 設定為 True，恢復自動觸發事件
        End If
End Sub
