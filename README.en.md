# KExcelAPI [![Build Status](https://travis-ci.org/webarata3/KExcelAPI.svg?branch=master)](https://travis-ci.org/webarata3/KExcelAPI) [![Coverage Status](https://coveralls.io/repos/webarata3/KExcelAPI/badge.svg?branch=master&service=github)](https://coveralls.io/github/webarata3/KExcelAPI?branch=master)

It is a wrapper of Apache POI for Kotlin. I am thinking of making everyone love Excel as easy as possible.

This library has been created under the influence of [GExcelAPI](https://github.com/nobeans/gexcelapi).

## How to use

From the sheet you can retrieve the Cell object by specifying the cell name or cell index.

Since we need a type to retrieve the data from the cell object, we will get the value with methods like toInt, toDouble, toStr etc.

In the case of a set of data, the type is obvious, so you can set the value on Excel with "=".

As a basic method of use, we import the `link.webarata3.kexcelapi. *` Package because we are using extension functions.

```kotlin
import link.webarata3.kexcelapi.*
```

Preparation is complete by this alone. You can now access Excel as follows.

```kotlin
// Easy file open, close
KExcel.open("file/book1.xlsx").use { workbook ->
    val sheet = workbook[0]

    // Cell loading
    // Access by Cell name
    println("""B7=${sheet["B7"].toStr()}""")
    // Access by Cell index [x, y]
    println("B7=${sheet[1, 6].toDouble()}")
    println("B7=${sheet[1, 6].toInt()}")

    // Write Cell
    sheet["A1"] = "あいうえお" // Japanese
    sheet[3, 7] = 123

    // Easy to write files
    KExcel.write(workbook, "file/book2.xlsx")
}
```

## Maven

The Maven repository (Gradle setting) is as follows.

```groovy
repositories {
    maven { url 'https://raw.githubusercontent.com/webarata3/maven/master/repository' }
}

dependencies {
    compile 'link.arata.kexcelapi:kexcelapi:0.1.0'
}
```

## License
MIT
