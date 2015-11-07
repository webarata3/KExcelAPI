package link.arata.kexcelapi

import org.hamcrest.Matchers.closeTo
import org.junit.Assert.assertThat
import org.junit.BeforeClass
import org.junit.Test
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat
import kotlin.test.fail
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

            assertThat(sheet("A1").getString(), IS("あ"))
            assertThat(sheet("B2").getString(), IS("い"))
            assertThat(sheet("C3").getString(), IS("う"))

            assertThat(sheet("A6").getInt(), IS(1))
            assertThat(sheet("B6").getInt(), IS(2))
            assertThat(sheet("C6").getDouble(), IS(3.0))

            assertThat(sheet("A7").getDouble(), IS(1.5))
            assertThat(sheet("B7").getDouble(), IS(2.5))
            assertThat(sheet("C7").getInt(), IS(3))

            val sdf = SimpleDateFormat("yyyy/MM/dd HH:mm")
            assertThat(sdf.format(sheet("A8").getDate()), IS("2015/01/03 08:15"))
            assertThat(sdf.format(sheet("B8").getDate()), IS("1899/12/31 11:27"))
            assertThat(sdf.format(sheet("C8").getDate()), IS("2015/10/01 00:00"))
        }
    }

    @Test
    fun test() {
        fail()
    }

    @Test
    fun セルのインデックスでの読み込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet(0, 0).getString(), IS("あ"))
            assertThat(sheet(1, 1).getString(), IS("い"))
            assertThat(sheet(2, 2).getString(), IS("う"))

            assertThat(sheet(0, 5).getInt(), IS(1))
            assertThat(sheet(1, 5).getInt(), IS(2))
            assertThat(sheet(2, 5).getDouble(), IS(3.0))

            assertThat(sheet(0, 6).getDouble(), IS(1.5))
            assertThat(sheet(1, 6).getDouble(), IS(2.5))
            assertThat(sheet(2, 6).getInt(), IS(3))

            val sdf = SimpleDateFormat("yyyy/MM/dd HH:mm")
            assertThat(sdf.format(sheet(0, 7).getDate()), IS("2015/01/03 08:15"))
            assertThat(sdf.format(sheet(1, 7).getDate()), IS("1899/12/31 11:27"))
            assertThat(sdf.format(sheet(2, 7).getDate()), IS("2015/10/01 00:00"))
        }
    }

    @Test
    fun 計算式の文字列取得テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet("A1").getString(), IS("a"))
            assertThat(sheet("B1").getString(), IS("3.0"))
            assertThat(sheet("C1").getString(), IS("true"))
            assertThat(sheet("D1").getString(), IS(""))
        }
    }

    @Test
    fun 計算式の中のInt型取得テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet("A2").getInt(), IS(33))
            assertThat(sheet("B2").getInt(), IS(5))
        }
    }

    @Test
    fun 計算式の中のInt型取得の小数の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet("A3").getInt(), IS(44))
            assertThat(sheet("B3").getInt(), IS(33))
        }
    }

    @Test
    fun 計算式の中のDouble型取得の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            assertThat(sheet("A3").getDouble(), IS(closeTo(44.5, 44.5)))
            assertThat(sheet("B3").getDouble(), IS(closeTo(33.29, 33.31)))
        }
    }

    @Test
    fun 計算式の中のDate型取得の場合テスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(1)

            val sdf = SimpleDateFormat("yyyy/MM/dd")
            assertThat(sdf.format(sheet("B4").getDate()), IS("2015/05/01"))
        }
    }

    @Test
    fun セルのラベルでの書き込みテスト() {
        KExcel.open("$BASE_DIR/book1.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            sheet("A1") value 100
            sheet("A2") value "あいうえお"

            KExcel.write(workbook, "$BASE_DIR/book2.xlsx")
        }

        KExcel.open("$BASE_DIR/book2.xlsx").use { workbook ->
            val sheet = workbook.getSheetAt(0)

            assertThat(sheet("A1").getInt(), IS(100))
            assertThat(sheet("A2").getString(), IS("あいうえお"))
        }
        Files.delete(Paths.get("$BASE_DIR/book2.xlsx"))
    }
}
