# ExplorerSingleMode

## 目的

Windows11でタブ化したエクスプローラをシングルウィンドウで運用します。  
以下の機能があります。
- ExplorerSingleMode起動時にエクスプローラが複数開いていた場合、どれかのエクスプローラウィンドウに他のエクスプローラタブが集約される
- 新しくエクスプローラを開いた場合は、新しいエクスプローラのタブを既に開いているウィンドウへ移動

## 動作環境

- Windows11 22H2以降
- .NET 6.0

  動作確認はx64環境でしか行っていません。

## 利用条件

このソフトウェアはMITライセンスの元で公開されたオープンソースソフトウェアです。  
利用する上で対価不要なためフリーソフトウェアとして扱って問題ありませんが、このソフトウェアを使用して発生したいかなる損害についても責任を負いません。  
ライセンスの全文は[こちら](https://opensource.org/licenses/mit-license.php)を参照してください。

## 使用方法

- インストール
  1. 適当な場所にフォルダを作成し、ZIPファイルの中身を展開します。
  1. `TaskRegist.cmd` を管理者として実行します。
- アンインストール  
  1. ZIPファイルの中身を展開したフォルダをエクスプローラで開きます。
  1. `TaskRemove.cmd` を管理者として実行します。
  1. ZIPファイルの中身を展開したフォルダを削除します。

## トラブルシューティング

- 新しいエクスプローラを開いたら、ウィンドウが移動するだけでタブが既にあるウィンドウに移ってくれない  
  ExplorerSingleMode による操作の指示に、Windowsの反応が追い付いていない可能性があります。  
  以下の方法を試してください。
  1. 「タスクスケジューラ」を開きます。
  1. 「タスク スケジューラ ライブラリ」から「ExplorerSingleMode」を探し、右クリックメニューから「終了」をクリックします。
  1. 「ExplorerSingleMode」を右クリックし、「プロパティ(P)」をクリックします。
  1. プロパティ画面が開きますので、「操作」タブをクリックします。
  1. 「編集」ボタンをクリックします。
  1. 「操作の編集」ダイアログが開きますので、「引数の追加(オプション)(A):」に数字を入れます。  
     この数字は、ExplorerSingleMode が行う操作と操作の間に挟む時間を調整するもので、大きい数字を入れるとそのぶん時間がかかるようになります。  
     単位は msec で、1つのタブあたり 入力値 * 5 (msec) だけ余計に時間がかかります。  
     目安としては、50から始めてダメなら50刻みで増やしつつ様子を見るのが良いと思います。
  1. 「操作の編集」ダイアログとプロパティ画面で「OK」を押して「タスクスケジューラ」へ戻ります。
  1. 「ExplorerSingleMode」を右クリックし、「実行する(R)」をクリックします。
     状態が「実行中」になったら完了です。
     再度エクスプローラを開いてみて、上手くいくよう数値を調整してみましょう。

## TODO

- 操作デモ

## 作者連絡先

Mail : [oss.develop.public@hosiyama.net](<mailto:oss.develop.public@hosiyama.net>)

Copyright © 2023 Hitoshi Hoshiyama All Rights Reserved.  
This project is licensed under the MIT License, see the [this site](https://opensource.org/licenses/mit-license.php).
