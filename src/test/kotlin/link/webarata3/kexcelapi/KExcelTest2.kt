package link.webarata3.kexcelapi

import org.apache.poi.ss.usermodel.Workbook
import org.hamcrest.Matchers.*
import org.junit.Assert.assertThat
import org.junit.Rule
import org.junit.Test
import org.junit.experimental.runners.Enclosed
import org.junit.experimental.theories.DataPoints
import org.junit.experimental.theories.Theories
import org.junit.experimental.theories.Theory
import org.junit.rules.ExpectedException
import org.junit.rules.TemporaryFolder
import org.junit.runner.RunWith
import java.io.FileNotFoundException
import java.util.*

@RunWith(Enclosed::class)
class KExcelTest2 {
    class 正常系_open {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        @Throws(Exception::class)
        fun openFileNameTest() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            val wb = KExcel.open(file.canonicalPath)
            assertThat<Workbook>(wb, `is`(notNullValue()))
            wb.close()
        }
    }

    class 異常系_open {
        @Test(expected = FileNotFoundException::class)
        fun openFileNameTest() {
            KExcel.open("/dummy")
        }
    }

    @RunWith(Theories::class)
    class 正常系_cellIndexToCellLabelTest {
        class Fixture(val x: Int, val y: Int, val cellLabel: String) {
            override fun toString(): String = "Fixture{x=$x, y=$y, cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            assertThat(fixture.toString(), KExcel.cellIndexToCellLabel(fixture.x, fixture.y), `is`(fixture.cellLabel))
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture(0, 0, "A1"),
                Fixture(1, 0, "B1"),
                Fixture(2, 0, "C1"),
                Fixture(26, 0, "AA1"),
                Fixture(27, 0, "AB1"),
                Fixture(28, 0, "AC1")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_セルのラベルでの読み込みテスト {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B2", "あいうえお"),
                Fixture("C3", "123"),
                Fixture("D4", "150.51"),
                Fixture("C2", "123")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_インデックスでのセルの読み込みテスト {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val x: Int, val y: Int, val expected: String) {
            override fun toString(): String = "Fixture{x=$x, y=$y, expected=$expected}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.x, fixture.y].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture(1, 1, "あいうえお"),
                Fixture(2, 2, "123"),
                Fixture(3, 3, "150.51"),
                Fixture(2, 1, "123")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_toStr {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B2", "あいうえお"),
                Fixture("C2", "123"),
                Fixture("D2", "150.51"),
                Fixture("F2", "true"),
                Fixture("G2", "123150.51"),
                Fixture("H2", ""),
                Fixture("I2", ""),
                Fixture("J2", "あいうえお123")
            )
        }
    }

    @RunWith(Theories::class)
    class 異常系_toStr {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        class Fixture(val cellLabel: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(UnsupportedOperationException::class.java)
                sheet[fixture.cellLabel].toStr()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("E2")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_toInt {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: Int) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toInt(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B3", 456),
                Fixture("C3", 123),
                Fixture("D3", 150),
                Fixture("G3", 369),
                Fixture("J3", 456123)
            )
        }
    }

    @RunWith(Theories::class)
    class 異常系_toInt {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        class Fixture(val cellLabel: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toInt()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B2"),
                Fixture("E3"),
                Fixture("F3"),
                Fixture("H3"),
                Fixture("I3"),
                Fixture("K3")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_toDouble {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: Double) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDouble(), `is`(closeTo(fixture.expected, 0.00001)))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B4", 123.456),
                Fixture("C4", 123.0),
                Fixture("D4", 150.51),
                Fixture("G4", 50.17),
                Fixture("J4", 123123.456)
            )
        }
    }

    @RunWith(Theories::class)
    class 異常系_toDouble {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        class Fixture(val cellLabel: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDouble()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B2"),
                Fixture("E4"),
                Fixture("F4"),
                Fixture("H4"),
                Fixture("I4"),
                Fixture("K4")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_toBoolean {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: Boolean) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toBoolean(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("F5", true),
                Fixture("G5", false)
            )
        }
    }

    @RunWith(Theories::class)
    class 異常系_toBoolean {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        class Fixture(val cellLabel: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toBoolean()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("B5"),
                Fixture("C5"),
                Fixture("D5"),
                Fixture("E5"),
                Fixture("K5")
            )
        }
    }

    @RunWith(Theories::class)
    class 正常系_toDate {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        class Fixture(val cellLabel: String, val expected: Date) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel, expected=$expected}"
        }

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDate(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("E6", TestUtil.getDate(2015, 12, 1)),
                Fixture("G6", TestUtil.getDate(2015, 12, 3))
            )
        }
    }

    @RunWith(Theories::class)
    class 異常系_toDate {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        class Fixture(val cellLabel: String) {
            override fun toString(): String = "Fixture{cellLabel=$cellLabel}"
        }

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDate()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                Fixture("A6"),
                Fixture("B6"),
                Fixture("C6"),
                Fixture("D6"),
                Fixture("F6"),
                Fixture("H6"),
                Fixture("I6"),
                Fixture("J6"),
                Fixture("K6")
            )
        }
    }
}
