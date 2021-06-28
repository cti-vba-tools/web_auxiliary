# メール補助機能

## outlookMailUsingIE.clsの使い方

### Module
```
With New outlookMailUsingIE
  .Send "subject", "body", "to_address"
End With
```

### 最終的に作成したい形
```
With New outlookMailUsingIE
  .setMode Default
  ' .setMode Outlook
  ' .setMode Gmail
  .setSubject "件名"
  .addTo "test_to@gmail.com"
  .addCc "test_cc@gmail.com" ' Outlook Web版では使用不可
  .addBcc "test_bcc@gmail.com" ' Outlook Web版では使用不可
  
  ' パターン①　一行ごと設定
  .addBody "〇●"
  .addBody ""
  .addBody "お疲れ様です。"
  .addBody ""
  .addBody "このメールはテストメールです。"
  
  ' パターン②　配列で一括して設定
  .setBody Array("〇●","","お疲れ様です。","","このメールはテストメールです。")
  
  ' パターン③　セル指定で設定
  .setBody Range("A1:A10").value
  
  .Send
End With
```

### 参照サイト
- [https://qiita.com/yoshi_782/items/a5d0a3f7ef30f5a36962](https://qiita.com/yoshi_782/items/a5d0a3f7ef30f5a36962)

### 注意事項
『CC』や『BCC』の指定はoutlook側の仕様上できないです！

