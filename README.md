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
