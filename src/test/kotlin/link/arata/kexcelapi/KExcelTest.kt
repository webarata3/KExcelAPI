/*
 * Copyright 2015 Shinichi ARATA
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package link.arata.kexcelapi

import org.hamcrest.Matchers.closeTo
import org.junit.Assert.assertThat
import org.junit.BeforeClass
import org.junit.Test
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat
import org.hamcrest.Matchers.`is` as IS

class KExcelTest() {
    companion object {
        var BASE_DIR = ""

        @JvmStatic
        @BeforeClass
        fun BeforeClass() {
            val path = Paths.get(KExcelTest::class.java.getResource("book1.xlsx").path).parent
            BASE_DIR = path.toString()
        }
    }

    @Test
    fun セルのラベルでの読み込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet["A1"].toStr(), IS("あ"))
            assertThat(sheet["B2"].toStr(), IS("い"))
            assertThat(sheet["C3"].toStr(), IS("う"))

            assertThat(sheet["A6"].toInt(), IS(1))
            assertThat(sheet["B6"].toInt(), IS(2))
            assertThat(sheet["C6"].toDouble(), IS(3.0))

            assertThat(sheet["A7"].toDouble(), IS(1.5))
            assertThat(sheet["B7"].toDouble(), IS(2.5))
            assertThat(sheet["C7"].toInt(), IS(3))

            val sdf = SimpleDateFormat("yyyy/MM/dd HH:mm")
            assertThat(sdf.format(sheet["A8"].toDate()), IS("2015/01/03 08:15"))
            assertThat(sdf.format(sheet["B8"].toDate()), IS("1899/12/31 11:27"))
            assertThat(sdf.format(sheet["C8"].toDate()), IS("2015/10/01 00:00"))
        }
    }

    @Test
    fun セルのインデックスでの読み込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet[0, 0].toStr(), IS("あ"))
            assertThat(sheet[1, 1].toStr(), IS("い"))
            assertThat(sheet[2, 2].toStr(), IS("う"))

            assertThat(sheet[0, 5].toInt(), IS(1))
            assertThat(sheet[1, 5].toInt(), IS(2))
            assertThat(sheet[2, 5].toDouble(), IS(3.0))

            assertThat(sheet[0, 6].toDouble(), IS(1.5))
            assertThat(sheet[1, 6].toDouble(), IS(2.5))
            assertThat(sheet[2, 6].toInt(), IS(3))

            val sdf = SimpleDateFormat("yyyy/MM/dd HH:mm")
            assertThat(sdf.format(sheet[0, 7].toDate()), IS("2015/01/03 08:15"))
            assertThat(sdf.format(sheet[1, 7].toDate()), IS("1899/12/31 11:27"))
            assertThat(sdf.format(sheet[2, 7].toDate()), IS("2015/10/01 00:00"))
        }
    }

    @Test
    fun 同じセルに違う方法で2回アクセス() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet["B2"].toStr(), IS("い"))
            assertThat(sheet[1, 1].toStr(), IS("い"))
            assertThat(sheet["B7"].toStr(), IS("2.5"))
            assertThat(sheet[1, 6].toStr(), IS("2.5"))
        }
    }

    @Test
    fun 計算式の文字列取得テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet["A1"].toStr(), IS("a"))
            assertThat(sheet["B1"].toStr(), IS("3.0"))
            assertThat(sheet["C1"].toStr(), IS("true"))
            assertThat(sheet["D1"].toStr(), IS(""))
        }
    }

    @Test
    fun 計算式の中のInt型取得テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet["A2"].toInt(), IS(33))
            assertThat(sheet["B2"].toInt(), IS(5))
        }
    }

    @Test
    fun 計算式の中のInt型取得の小数の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet["A3"].toInt(), IS(44))
            assertThat(sheet["B3"].toInt(), IS(33))
        }
    }

    @Test
    fun 計算式の中のDouble型取得の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet["A3"].toDouble(), IS(closeTo(44.5, 44.5)))
            assertThat(sheet["B3"].toDouble(), IS(closeTo(33.29, 33.31)))
        }
    }

    @Test
    fun 計算式の中のDate型取得の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            val sdf = SimpleDateFormat("yyyy/MM/dd")
            assertThat(sdf.format(sheet["B4"].toDate()), IS("2015/05/01"))
        }
    }

    @Test
    fun セルのラベルでの書き込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            sheet["A1"].setValue(100)
            sheet["A2"].setValue("あいうえお")

            KExcel.write(workbook, "$BASE_DIR/book2.xlsx")
        }

        KExcel.open("$BASE_DIR/book2.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet["A1"].toInt(), IS(100))
            assertThat(sheet["A2"].toStr(), IS("あいうえお"))
        }
        Files.delete(Paths.get("$BASE_DIR/book2.xlsx"))
    }

    @Test
    fun シートのラベルからの書き込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            sheet["A1"] = 100
            sheet["A2"] = "あいうえお"

            KExcel.write(workbook, "$BASE_DIR/book2.xlsx")
        }

        KExcel.open("$BASE_DIR/book2.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet["A1"].toInt(), IS(100))
            assertThat(sheet["A2"].toStr(), IS("あいうえお"))
        }
        Files.delete(Paths.get("$BASE_DIR/book2.xlsx"))
    }

    @Test
    fun シートのインデックスからの書き込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            sheet[0, 0] = 100
            sheet[0, 1] = "あいうえお"

            KExcel.write(workbook, "$BASE_DIR/book2.xlsx")
        }

        KExcel.open("$BASE_DIR/book2.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet["A1"].toInt(), IS(100))
            assertThat(sheet["A2"].toStr(), IS("あいうえお"))
        }
        Files.delete(Paths.get("$BASE_DIR/book2.xlsx"))
    }
}
