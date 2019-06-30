# extract_image_from_MSword

MS Wordファイルから画像を抽出します。

## 処理概要

1. main
   1. `AppSettings.json`から設定情報を取得
   1. wordファイルを画像格納ディレクトリにコピー
   1. wordファイルの拡張子をzipに変更
   1. zipファイルを解凍する。
   1. `word\media`に画像ファイルが入っているので、画像ファイルを解凍したディレクトリ直下に移動
   1. 画像ファイル以外を削除
   1. 画像ファイルが50個以上ある場合、サブディレクトリを作成し（01、02…）、50個ずつに分割して格納
   1. 画像ファイルの名前を変更
      1. `image1.png`の`image`を削除
      1. 数字が1桁の場合は、先頭に0を付与

## 設定ファイル

JSONファイルで行う。  
プロジェクトルートに`appsettings.json`を作成する。  
ファイルは出力ディレクトリにコピーする設定にする。

```json
{
  "AppSettings": {
    "TargetDir": "D:\\work\\wordImage_word",
    "ImageDir": "D:\\work\\wordImage_image"
  }
}
```

* TargetDir: 抽出対象のwordファイルを格納しているディレクトリ
* TargImageDiretDirs: wordから抽出した画像ファイルを格納するディレクトリ
