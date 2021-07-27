# README

ジョブカンの工数管理の入力をエクセルのデータを元に自動でします。

## 環境準備

NODEバージョン： `.node-version` をご確認ください。

以下のコマンドを実行

```
brew tap homebrew/cask
brew install chromedriver --cask
yarn install
```

sample.xlsxをコピーして「工数入力.xlsx」にリネームする。


## 使い方

工数入力.xlsxを開いて登録内容を編集する。
登録するデータのシートを指定するので過去の内容も保存しておきたい場合はコピーして新しいシートを増やして登録することも可能です。

以下のコマンドを実行する。

ドライランモードでテストしてから本登録モードをすることをオススメします。

間違った内容が登録されたら後で消すのが大変です。

### コマンド

シート名：工数入力.xlsxで登録するシートの名前を設定する
アカウント名：ジョブカンのアカウント名
パスワード：ジョブカンのパスワード

#### ドライランモード

実際に登録はしません。

```
yarn dryrun --sheet シート名 --account アカウント名 --password パスワード
```

#### 本登録モード

実際に登録します。

```
yarn start --sheet シート名 --account アカウント名 --password パスワード
```
