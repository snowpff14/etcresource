# 使い方
## 導入
* pythonをインストール(3.7で確認)
* 以下のコマンドでライブラリのインストールを行う
* pythonをインストールできないなどの場合はexeの方を使うこと
```
pip install pandas
```

## 使い方

1. `resources/appConfig.ini`に各種パスを設定しているので適宜修正すること以下はデフォルトの設定値での説明
1. マージしたいCSVファイルを`input`に格納する。この中にあるファイルをすべてマージするため不要なファイルを削除しておくこと
1. `fileMerge.exe`を実行する。pythonをインストールして動かしたいときはfileMerge.pyを実行する。
1. `output`にマージしたファイルが出力される。


## ソースの説明

1. 指定したディレクトリ下から`.csv`のファイルを取得する。
1. 取得したファイルの1つ目のヘッダーをもとにほかのファイルを順次結合する。
1. 結合後のファイルをダブルクォーテーションくくりで出力
