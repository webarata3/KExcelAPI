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
        fun open(fileName: String): Workbook {
            return WorkbookFactory.create(FileInputStream(Paths.get(fileName).toFile()))
        }

        @JvmStatic
        fun write(workbook: Workbook, fileName: String) {
            var outputPath = Paths.get(fileName)
            try {
                Files.newOutputStream(outputPath).use {
                    workbook.write(it)
                }
            } catch (e: IOException) {
                e.printStackTrace()
            }
        }
    }
}

public operator fun Sheet.invoke(n: Int): Row {
    return getRow(n) ?: createRow(n)
}

public operator fun Row.invoke(n: Int): Cell {
    return getCell(n) ?: createCell(n, Cell.CELL_TYPE_BLANK)
}

public operator fun Sheet.invoke(x: Int, y: Int): Cell {
    var row = this(y)
    return row(x)
}

private val ORIGIN = 'A'.toInt()
private val RADIX = 26

// https://github.com/nobeans/gexcelapi/blob/master/src/main/groovy/org/jggug/kobo/gexcelapi/GExcel.groovy
public operator fun Sheet.invoke(cellLabel: String): Cell {
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
    return this(num, matcher.group(2).toInt() - 1)
}

public fun Cell.getString(): String {
    when (cellType) {
        Cell.CELL_TYPE_STRING -> return stringCellValue
        Cell.CELL_TYPE_NUMERIC -> return numericCellValue.toString()
        Cell.CELL_TYPE_BOOLEAN -> return booleanCellValue.toString()
        Cell.CELL_TYPE_BLANK -> return ""
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING -> return cellValue.stringValue
                Cell.CELL_TYPE_NUMERIC -> return cellValue.numberValue.toString()
                Cell.CELL_TYPE_BOOLEAN -> return cellValue.booleanValue.toString()
                else -> throw IllegalAccessException("cellはStringに変換できません")
            }
        }
        else -> throw IllegalAccessException("cellはStringに変換できません")
    }
}

public fun Cell.getInt(): Int {
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

public fun Cell.getDouble(): Double {
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

public fun Cell.getDate(): Date {
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

public infix fun Cell.value(value: Any) {
    when (value) {
        is String -> setCellValue(value)
        is Int -> setCellValue(value.toDouble())
        is Double -> setCellValue(value)
        else -> throw IllegalArgumentException("文字列か数値のみ対応しています")
    }
}
