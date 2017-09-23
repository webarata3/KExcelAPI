/*
The MIT License (MIT)

Copyright (c) 2017 Shinichi ARATA.

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
package link.webarata3.kexcelapi

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.CellValue
import org.apache.poi.ss.usermodel.DateUtil
import java.util.*

class CellProxy(private val cell: Cell) {
    private var cellValue: CellValue? = null

    init {
        if (cell.cellTypeEnum == CellType.FORMULA) {
            cellValue = getFomulaCellValue(cell)
        }
    }

    private fun getCellTypeEnum(): CellType {
        return if (cellValue == null) {
            cell.cellTypeEnum
        } else {
            (cellValue as CellValue).cellTypeEnum
        }
    }

    private fun getStringCellValue(): String {
        return if (cellValue == null) cell.stringCellValue else (cellValue as CellValue).stringValue
    }

    private fun getNumericCellValue(): Double {
        return if (cellValue == null) cell.numericCellValue else (cellValue as CellValue).numberValue
    }

    private fun getBooleanCellValue(): Boolean {
        return if (cellValue == null) cell.booleanCellValue else (cellValue as CellValue).booleanValue
    }

    private fun isDateType(): Boolean {
        return if (cellValue == null) {
            if (cell.cellTypeEnum == CellType.NUMERIC) DateUtil.isCellDateFormatted(cell)
            else false
        } else {
            if ((cellValue as CellValue).cellTypeEnum == CellType.NUMERIC) DateUtil.isCellDateFormatted(cell)
            else false
        }
    }

    private fun normalizeNumericString(numeric: Double): String {
        // 44.0のような数値を44として取得するために、入力された数値と小数点以下を切り捨てた数値が
        // 一致した場合には、intにキャストして、小数点以下が表示されないようにしている
        return if (numeric == Math.ceil(numeric)) {
            numeric.toInt().toString()
        } else numeric.toString()
    }

    private fun stringToInt(value: String): Int {
        try {
            return java.lang.Double.parseDouble(value).toInt()
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cellはintに変換できません")
        }
    }

    private fun stringToDouble(value: String): Double {
        try {
            return java.lang.Double.parseDouble(value)
        } catch (e: NumberFormatException) {
            throw IllegalAccessException("cellはdoubleに変換できません")
        }
    }

    private fun getFomulaCellValue(cell: Cell): CellValue {
        val wb = cell.sheet.workbook
        val helper = wb.creationHelper
        val evaluator = helper.createFormulaEvaluator()
        return evaluator.evaluate(cell)
    }

    fun toStr(): String {
        when (getCellTypeEnum()) {
            CellType.STRING -> return getStringCellValue()
            CellType.NUMERIC -> return if (isDateType()) {
                throw UnsupportedOperationException("今はサポート外")
            } else {
                normalizeNumericString(getNumericCellValue())
            }
            CellType.BOOLEAN -> return getBooleanCellValue().toString()
            CellType.BLANK -> return ""
            else // _NONE, ERROR
            -> throw IllegalAccessException("cellはStringに変換できません")
        }
    }

    fun toInt(): Int {
        when (getCellTypeEnum()) {
            CellType.STRING -> return stringToInt(getStringCellValue())
            CellType.NUMERIC -> return if (isDateType()) {
                throw IllegalAccessException("cellはIntに変換できません")
            } else {
                getNumericCellValue().toInt()
            }
            else -> throw IllegalAccessException("cellはIntに変換できません")
        }
    }

    fun toDouble(): Double {
        when (getCellTypeEnum()) {
            CellType.STRING -> return stringToDouble(getStringCellValue())
            CellType.NUMERIC -> return if (isDateType()) {
                throw IllegalAccessException("cellはDoubleに変換できません")
            } else {
                getNumericCellValue()
            }
            else -> throw IllegalAccessException("cellはDoubleに変換できません")
        }
    }

    fun toBoolean(): Boolean {
        when (getCellTypeEnum()) {
            CellType.BOOLEAN -> return getBooleanCellValue()
            else -> throw IllegalAccessException("cellはBooleanに変換できません")
        }
    }

    fun toDate(): Date {
        when {
            isDateType() -> return cell.dateCellValue
            else -> throw IllegalAccessException("cellはDateに変換できません")
        }
    }
}
