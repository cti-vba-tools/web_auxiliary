# メール補助機能

## 導入方法について
- [【ExcelVBA入門】クラスモジュールのメリット・使い方を徹底解説！](https://www.sejuku.net/blog/67476)

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
  .addCc "test_cc@gmail.com" ' Outlook Web版では使用不可のためToに結合
  .addBcc "test_bcc@gmail.com" ' Outlook Web版では使用不可のためToに結合
  
  .addTemplate "担当者", Range("C2").value
  
  ' パターン①　一行ごと設定
  .addBody "{{担当者}}さま"
  .addBody ""
  .addBody "お疲れ様です。"
  .addBody ""
  .addBody "このメールはテストメールです。"
  
  ' パターン②　配列で一括して設定
  .setBody Array("{{担当者}}さま","","お疲れ様です。","","このメールはテストメールです。")
  
  ' パターン③　セル指定で設定
  .setBody Range("A1:A10").value
  
  .Preview
End With
```

### 参照サイト
- [APIを使わずにVBAでGmailを作成](https://qiita.com/yoshi_782/items/a5d0a3f7ef30f5a36962)

### 注意事項
『CC』や『BCC』の指定はoutlook側の仕様上できないです！

