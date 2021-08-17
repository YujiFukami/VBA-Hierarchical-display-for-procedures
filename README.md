# VBA-Hierarchical-display-for-procedures
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

# 使い方

## Step1
Excelの設定でExcel2019の場合「Excelのオプション」→「トラストセンター」→「マクロの設定」で

「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れておく
![Excelの設定](https://user-images.githubusercontent.com/73621859/126287884-57db4a75-3f34-4b35-b23d-f705067a1869.jpg)

## Step2
「frmKaiso.fm」「ModExtProcedure.bas」「ClassVBProject」「ClassModule」「ClassProcedure」をダウンロードし、VBEにインポートする。
また、下記ライブラリを追加で参照すること。

- 「Microsoft Forms 2.0 Object Library」→ListView,TreeViewを動かすためっぽい

- 「Microsoft Windows Common Controls 6.0(SP6)」→ListView,TreeViewを動かすためっぽい

- 「Microsoft Visual Basic for Applications Extensibility 5.3」→VBAコードをVBAで参照するため

![階層化フォーム 参照ライブラリ](https://user-images.githubusercontent.com/73621859/128787617-59d52e7e-0439-4f6c-9877-4bfe11e8d745.jpg)

TreeViewコントロールはExcelバージョン,Windows環境で動いたり動かなかったりするらしいので注意すること。

実行環境など報告していただくと感謝感激雨霰。


## Step3
セルに「=Kaiso()」と入力するとプロシージャの一覧、階層化表示のフォームが出現する。
![1 KAISO()](https://user-images.githubusercontent.com/73621859/126260383-018720ef-904d-48ed-a82c-41041c497c89.jpg)

# 階層化表示フォームの画面および使い方説明
![階層フォーム説明1](https://user-images.githubusercontent.com/73621859/128684001-6fba88ef-dc7f-4ec6-bf7d-f79c0692b225.jpg)

![階層フォーム説明2](https://user-images.githubusercontent.com/73621859/128684028-3413017b-b556-4c15-b247-87dbd582f6e8.jpg)

# 使用デモ動画
![階層化実行テスト](https://user-images.githubusercontent.com/73621859/128684086-2a0e3bdd-f528-48b0-b148-f86db97ca655.gif)

