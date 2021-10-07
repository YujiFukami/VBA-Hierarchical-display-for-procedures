# VBA-Hierarchical-display-for-procedures
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

実行環境など報告していただくと感謝感激雨霰。

# 使い方

下記ステップを踏むか、「Sample Kaiso.xlsm」内の設定を利用してください。

## Step1
Excelの設定でExcel2019の場合「Excelのオプション」→「トラストセンター」→「マクロの設定」で

「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れておく
![Excelの設定](https://user-images.githubusercontent.com/73621859/126287884-57db4a75-3f34-4b35-b23d-f705067a1869.jpg)

設定方法の記事↓
http://blog.livedoor.jp/aero_iki-jibundakemacro/archives/30630575.html

## Step2
下記ダウンロードし、VBEにインポートする。
・frmKaiso.fm
・ModExtProcedure.bas
・ModGetProcedureAllCode.bas
・ModRefLibraryForKaiso.bas
・ModUserFormResize.bas
・ClassVBProject
・ClassModule
・ClassProcedure

また、下記ライブラリを追加で参照する、もしくは「ModRefLibraryForKaiso.bas」モジュール内の「RefLibraryForKaiso」プロシージャを実行すること。

- 「Microsoft Forms 2.0 Object Library」→ListView,TreeViewを動かすためっぽい

- 「Microsoft Windows Common Controls 6.0(SP6)」→ListView,TreeViewを動かすためっぽい

- 「Microsoft Visual Basic for Applications Extensibility 5.3」→VBAコードをVBAで参照するため

![階層化フォーム 参照ライブラリ](https://user-images.githubusercontent.com/73621859/128787617-59d52e7e-0439-4f6c-9877-4bfe11e8d745.jpg)

TreeViewコントロールはExcelバージョン,Windows環境で動いたり動かなかったりするらしいので注意すること。


## Step3
セルに「=Kaiso()」と入力するとプロシージャの一覧、階層化表示のフォームが出現する。
![1 KAISO()](https://user-images.githubusercontent.com/73621859/126260383-018720ef-904d-48ed-a82c-41041c497c89.jpg)

# 階層化表示フォームの画面および使い方説明
![階層フォーム説明1](https://user-images.githubusercontent.com/73621859/128684001-6fba88ef-dc7f-4ec6-bf7d-f79c0692b225.jpg)

![階層フォーム説明2](https://user-images.githubusercontent.com/73621859/128684028-3413017b-b556-4c15-b247-87dbd582f6e8.jpg)

# 外部参照プロシージャの一覧表示の仕組みの説明

個人的に使用していたりする「開発用アドイン」を参照して、新規にマクロ付ブックを開発した後、「開発用アドイン」で参照しているプロシージャを一覧で取得し「新規開発ブック」上にコードをコピーすることができる。
コードのコピーののちに、「開発用アドイン」の参照を解除することができる。

![外部参照プロシージャ一覧取得の概要図](https://user-images.githubusercontent.com/73621859/131796576-9489b7d6-f7d0-4af8-8345-eb380cd35731.jpg)

# 使用デモ動画
![階層化実行テスト](https://user-images.githubusercontent.com/73621859/128684086-2a0e3bdd-f528-48b0-b148-f86db97ca655.gif)


# 謝辞
本プログラム開発にご協力いただいた方々。
協力ありがとうございます😆😆😆😆😆😆

https://github.com/furyutei
