/*
The MIT License (MIT)

Copyright (c) 2016 Shinichi ARATA.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
 */
package link.arata.kexcelapi

import org.apache.poi.ss.usermodel.*
import java.io.FileInputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths
import java.util.*
import java.util.regex.Pattern

public class KExcel {
    companion object {
        @JvmStatic
        public fun open(fileName: String): Workbook {
            return WorkbookFactory.create(FileInputStream(Paths.get(fileName).toFile()))
        }

        @JvmStatic
        public fun write(workbook: Workbook, fileName: String) {
            var outputPath = Paths.get(fileName)
            try {
                Files.newOutputStream(outputPath).use {
                    workbook.write(it)
                }
            } catch (e: IOException) {
                throw e
            }
        }

        @JvmStatic
        public fun cellIndexToCellName(x: Int, y: Int): String {
            var cellName = dec26(x, 0)
            return cellName + (y + 1)
        }

        @JvmStatic
        private fun dec26(num: Int, first: Int): String {
            return if (num > 25) {
                dec26(num / 26, 1)
            } else {
                ""
            } + ('A' + (num - first) % 26)
        }
    }
}

public operator fun Workbook.get(n: Int): Sheet {
    return this.getSheetAt(n)
}

public operator fun Workbook.get(name: String): Sheet {
    return this.getSheet(name)
}

public operator fun Sheet.get(n: Int): Row {
    return getRow(n) ?: createRow(n)
}

public operator fun Row.get(n: Int): Cell {
    return getCell(n) ?: createCell(n, Cell.CELL_TYPE_BLANK)
}

public operator fun Sheet.get(x: Int, y: Int): Cell {
    var row = this[y]
    return row[x]
}

private val ORIGIN = 'A'.toInt()
private val RADIX = 26

// https://github.com/nobeans/gexcelapi/blob/master/src/main/groovy/org/jggug/kobo/gexcelapi/GExcel.groovy
public operator fun Sheet.get(cellLabel: String): Cell {
    val p1 = Pattern.compile("([a-zA-Z]+)([0-9]+)");
    val matcher = p1.matcher(cellLabel)
    matcher.find()

    var num = 0
    matcher.group(1).toUpperCase().reversed().forEachIndexed {
        i, c ->
        var delta = c.toInt() - ORIGIN + 1
        num += delta * Math.pow(RADIX.toDouble(), i.toDouble()).toInt()
    }
    num -= 1
    return this[num, matcher.group(2).toInt() - 1]
}

private fun normalizeNumericString(numeric: Double): String {
    return if (numeric == Math.ceil(numeric)) {
        numeric.toInt().toString()
    } else {
        numeric.toString()
    }
}

public fun Cell.toStr(): String {
    when (cellType) {
        Cell.CELL_TYPE_STRING -> return stringCellValue
        Cell.CELL_TYPE_NUMERIC -> return normalizeNumericString(numericCellValue)
        Cell.CELL_TYPE_BOOLEAN -> return booleanCellValue.toString()
        Cell.CELL_TYPE_BLANK -> return ""
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING -> return cellValue.stringValue
                Cell.CELL_TYPE_NUMERIC -> return normalizeNumericString(cellValue.numberValue)
                Cell.CELL_TYPE_BOOLEAN -> return cellValue.booleanValue.toString()
                Cell.CELL_TYPE_BLANK -> return ""
                else -> throw IllegalAccessException("cellはStringに変換できません")
            }

        }
        else -> throw IllegalAccessException("cellはStringに変換できません")
    }
}

public fun Cell.toInt(): Int {
    fun stringToInt(value: String): Int {
        try {
            // toIntだと44.5のような文字列を44に変換できないため、一度Dobuleに変換している
            return value.toDouble().toInt()
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cellはIntに変換できません")
        }
    }

    when (cellType) {
        Cell.CELL_TYPE_STRING -> return stringToInt(stringCellValue)
        Cell.CELL_TYPE_NUMERIC -> return numericCellValue.toInt()
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING -> return stringToInt(cellValue.stringValue)
                Cell.CELL_TYPE_NUMERIC -> return cellValue.numberValue.toInt()
                else -> throw IllegalAccessException("cellはIntに変換できません")
            }
        }
        else -> throw IllegalAccessException("cellはIntに変換できません")
    }
}

public fun Cell.toDouble(): Double {
    fun stringToDouble(value: String): Double {
        try {
            return value.toDouble()
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cellはDoubleに変換できません")
        }
    }

    when (cellType) {
        Cell.CELL_TYPE_STRING -> return stringToDouble(stringCellValue)
        Cell.CELL_TYPE_NUMERIC -> return numericCellValue.toDouble()
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING -> return stringToDouble(cellValue.stringValue)
                Cell.CELL_TYPE_NUMERIC -> return cellValue.numberValue.toDouble()
                else -> throw IllegalAccessException("cellはDoubleに変換できません")
            }
        }
        else -> throw IllegalAccessException("cellはDoubleに変換できません")
    }
}

public fun Cell.toBoolean(): Boolean {
    when (cellType) {
        Cell.CELL_TYPE_BOOLEAN -> return booleanCellValue
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_BOOLEAN -> return cellValue.booleanValue
                else -> throw IllegalAccessException("cellはBooleanに変換できません")
            }
        }
        else -> throw IllegalAccessException("cellはBooleanに変換できません")
    }
}

public fun Cell.toDate(): Date {
    when (cellType) {
        Cell.CELL_TYPE_NUMERIC -> return dateCellValue
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_NUMERIC -> return dateCellValue
                else -> throw IllegalAccessException("cellはDeteに変換できません")
            }
        }
        else -> throw IllegalAccessException("cellはDateに変換できません")
    }
}

private fun getFormulaCellValue(cell: Cell): CellValue {
    val workbook = cell.sheet.workbook
    val helper = workbook.creationHelper
    val evaluator = helper.createFormulaEvaluator()
    return evaluator.evaluate(cell)
}

public operator fun Sheet.set(cellLabel: String, value: Any) {
    this[cellLabel].setValue(value)
}

public operator fun Sheet.set(x: Int, y: Int, value: Any) {
    this[x, y].setValue(value)
}

private fun Cell.setValue(value: Any) {
    when (value) {
        is String -> setCellValue(value)
        is Int -> setCellValue(value.toDouble())
        is Double -> setCellValue(value)
        is Date -> setCellValue(value)
        is Boolean -> setCellValue(value)
        else -> throw IllegalArgumentException("文字列か数値のみ対応しています")
    }
}
