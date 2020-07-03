# アイコンエディタ(IconEditor)
EXCELでアイコンを作成することが出来ます。

## 外観

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/411106/56939e83-c86c-7198-b4d2-595fbf48c4f1.png)


## 動作環境

WINDOWS上の EXCEL2010以降 **<FONT COLOR="RED">32bit 64bit</FONT>** いずれのバージョンのEXCELでも動作します。


## 紹介

[EXCELをアイコンエディタにするツール](https://qiita.com/takahasinaoki/items/f3f49ac12df0634268a6)


## 特徴

### (1) ImageMso

ImageMsoを指定して、Officeに標準で登録されているアイコンを読込むことが出来ます。<br>
透明色や半透明色は反映されます。



### (2) ファイル形式

PNG, BMP, ICO, JPG, GIFファイルを読込むことが出来ます。<br>	
透明色や半透明色は反映されます。

PNG, BMP, ICOファイルで保存することが出来ます。<br>
PNGファイルで保存すれば、透明色や半透明色は保存されます。



### (3) アイコンサイズ

アイコンのサイズに制約を設けていません。(16x16,24x20,32x32,64x64等々)<br>
すべてのシートのすべてのセルに対して、自由なサイズでアイコンを描画することが出来ます。

読込める画像のサイズにも制約を設けていません。(ただし、大きすぎると予期せぬ不具合の原因になります)<br>		
書式（このツールではセルの色）をあまりにも多用しすぎると、不具合を起こすようです。


### (4) 透明色・半透明色

透明色・半透明色(アルファチャンネル)に対応しています。



### (5) UNDO

直前のコマンドのUNDOが可能です。



### (6) 図形(オートシェイプ)

図形(オートシェイプ)からアイコンを作成することが出来ます。



### (7) クリップボード

クリップボードの画像からアイコンを作成することが出来ます。<br>
例えばWEB上のアルファチャンネルが設定された画像をコピーした場合は、透明色・半透明色も反映されます。

作成中のアイコンをクリップボードにコピーすることが出来ます。<br>
ビットマップ形式だけでなく、PNG形式でもコピーされるため、透明色・半透明色も反映されます。



### (8) サンプル表示

編集中のアイコンをリボンに表示することが出来ます。



### (9) 一括読込・一括保存

複数のアイコンを一括で読込むことも、保存することも出来ます。



### (10) 色を値で指定

アイコンの色を数値や演算で指定することも可能です。

そのための便利なセル関数も用意してあります。



## ソースコードについて
当ツールを作成するにあたり、いろいろ調べて実現した機能がありますので、以下のようなサンプルを探している方は、ソースの中に参考になるロジックがあると思います。

- EXCELでコピー中のセル範囲(Range)を取得する
- VBAでBMPファイルを保存する
- VBAでICOファイルを保存する
- ImageMsoから透明色を反映した画像を取得する
- VBAでPNG形式のクリップボードデータの読込み および 書込み を行う
- VBAでGDI+ のAPI(64bit)を実行する
- VBAでGDI+ を使ってアルファチャネルの画像を取り扱う
- リボンのボタンのアイコンを動的に変更する

## ダウンロード
 [こちら](https://github.com/takanaweb5/IconEditor/releases) からダウンロード可能です。<br>
 [IconEditor_v@@@.zip] ファイルをダウンロードしてください。※@@@はバージョン



## ご意見
バグや要望は[こちら](https://github.com/takanaweb5/IconEditor/issues)まで
![image](https://user-images.githubusercontent.com/50874513/85219654-b3465a80-b3e0-11ea-937e-89708e6af1b8.png)

## ライセンス
MITライセンスに準拠します<br>
Copyright (c) 2019 TAKAHASHI Naoki(JPN) 高橋
