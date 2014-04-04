spreadsheet2sql
===============

make insert SQL from Google SpreadSheet


## 使用例
```
#! /bin/bash

python ./spreadsheet2sql.py "hogehoge@gmail.com" "hogepassword" "huga_sheetname"

echo 'end'

```

-----

## 備考

- 対象のワークシートは、名前が 英数_-なもののみ。全角を使用したシートは定義や出力等に利用出来る。
- 各ワークシート名はテーブル名として使用する。
- 各ワークシートの1行目は列名として使用する。またアンダーバーは使用出来ない。
- 列名にアンダーバーを含む場合、予めハイフンに変換しておく。SQL出力時にアンダーバーに変換している。
- 作成したSQLを出力用ワークシートに書き込むことが出来る。出力用ワークシートの1行目1列目には項目名が必要（アンダーバーは使えない）
- アプリとして２段階認証を設定する必要がある。詳細は、[Pythonのgdataで認証が成功しない](http://mamansoft.net/blog/python%E3%81%AEgdata%E3%81%A7%E8%AA%8D%E8%A8%BC%E3%81%8C%E6%88%90%E5%8A%9F%E3%81%97%E3%81%AA%E3%81%84/)を参考。
-------



## 課題

- 出力時、一旦ワークシート内のデータを消しているが、うまく消えてくれない。そのため、ループで無理やり消している。
- python、がんばる；；