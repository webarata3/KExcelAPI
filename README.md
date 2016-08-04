# KExcelAPI [![Build Status](https://travis-ci.org/webarata3/KExcelAPI.svg?branch=master)](https://travis-ci.org/webarata3/KExcelAPI) [![Coverage Status](https://coveralls.io/repos/webarata3/KExcelAPI/badge.svg?branch=master&service=github)](https://coveralls.io/github/webarata3/KExcelAPI?branch=master)

[Japanese](/README.ja.md)

It is a wrapper for Apache Poi written in Kotlin.
We believe to be able to as easy as possible to access the Excel.

This library has been affected from x [GExcelAPI](https://github.com/nobeans/gexcelapi).

## How to use

By specifying the index of the cell name or cell from a sheet, you can get the Cell object.

Since the acquisition of data from the cell object type is needed, to get the value in the method of  `toInt` , `toDouble` ,  `toStr` , etc.
Because in the case of a set of data type is self-evident, you can set a value on the Excel in the "=".

The first step is to import the `link.arata.kexcelapi.*` package because you are using an extension function.

```kotlin
import link.arata.kexcelapi.*
```

It's ready to go. Now, you can access to Excel as follows.

```kotlin
// Easy file open, close
KExcel.open("file/book1.xlsx").use { workbook ->
    val sheet = workbook[0]

    // Reading of the cell
    // Access at the cell name
    println("""B7=${sheet["B7"].toStr()}""")
    // Access of the index of the cell [x, y]
    println("B7=${sheet[1, 6].toDouble()}")
    println("B7=${sheet[1, 6].toInt()}")

    // Writing of cell
    sheet["A1"] = "ABCDE"
    sheet[3, 7] = 123

    // Also easy writing of files
    KExcel.write(workbook, "file/book2.xlsx")
}
```

## Maven

Maven（or Gradle） repository is as follows.

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
