# 使い方
## 導入
* pythonをインストール(3.7で確認)
* 以下のコマンドでライブラリのインストールを行う
* pythonをインストールできないなどの場合はexeの方を使うこと
```
pip install pandas
pip install numpy
```

## 使い方

1. `resources/検索データ.xlsx`に抽出したい条件の文言を記載
1. `resources/appConfig.ini`に抽出時の方法(前方一致、後方一致など)を指定する。またソート方法も指定を行う。
    * 文字列が一致したものとマッチングしたものを抽出する。今回は一致するものを抽出するだけなので関係はないが、一致したものを置換するときなどは文字列の長い方でソートを掛けた方がよい
1. 対象とするファイルを`data`に格納する。この中にあるファイルをすべて処理するため不要なファイルを削除しておくこと
1. `regExtract.exe`を実行する。pythonをインストールして動かしたいときはregExtract.pyを実行する。
1. `output`に抽出した行をまとめたファイルが出力される。


## ソースの説明

1. 指定したディレクトリ下からファイルを取得する。
1. 検索データ.xlsxに記載されている内容をappConfig.iniの設定に応じてソートを行い、抽出条件できるように文字列を作成する。
1. 正規表現を使って処理対象のファイルから抽出を行う
