# KExcelAPI [![Build Status](https://travis-ci.org/webarata3/KExcelAPI.svg?branch=master)](https://travis-ci.org/webarata3/KExcelAPI) [![Coverage Status](https://coveralls.io/repos/webarata3/KExcelAPI/badge.svg?branch=master&service=github)](https://coveralls.io/github/webarata3/KExcelAPI?branch=master)

[English](/README.en.md)

[KExcelAPIの紹介サイト](https://kexcelapi.webarata3.link)です。

Kotlin用のApache POIのラッパーです。皆さん大好きなExcelをできるだけ簡単にアクセスできるように考えています。

このライブラリーは[GExcelAPI](https://github.com/nobeans/gexcelapi)から影響を受けて作成しています。

## 使い方

シートから、セル名かセルのインデックスを指定してCellオブジェクトを取得することができます。
セルオブジェクトからのデータの取得には型が必要なので、toInt、toDouble、toStr等のメソッドで値を取得します。
データのセットの場合には型が自明なため、「=」でExcel上に値をセットできます。

使い方の基本としては、拡張関数を使用しているため `link.arata.kexcelapi.*` パッケージをimportします。

```kotlin
import link.webarata3.kexcelapi.*
```

これだけで準備は完了です。これで、次のようにExcelへアクセスできます。

```kotlin
// 簡単にファイルオープン、クローズ
KExcel.open("file/book1.xlsx").use { workbook ->
    val sheet = workbook[0]

    // セルの読み込み
    // セル名でのアクセス
    println("""B7=${sheet["B7"].toStr()}""")
    // セルのインデックスでのアクセス [x, y]
    println("B7=${sheet[1, 6].toDouble()}")
    println("B7=${sheet[1, 6].toInt()}")

    // セルの書き込み
    sheet["A1"] = "あいうえお"
    sheet[3, 7] = 123

    // ファイルの書き込みも簡単に
    KExcel.write(workbook, "file/book2.xlsx")
}
```

## Maven

Mavenのリポジトリ（Gradleの設定）は次のとおりです。

```groovy
dependencies {
    compile 'link.webarata3.kexcelapi:kexcelapi:0.5.1'
}
```

## ライセンス
MIT
