package link.webarata3.kexcelapi

import org.junit.rules.TemporaryFolder

import java.io.File
import java.nio.file.Files
import java.util.*

object TestUtil {
    @Throws(Exception::class)
    fun getTempWorkbookFile(tempFolder: TemporaryFolder, fileName: String): File {
        val tempFile = File(tempFolder.root, "temp.xlsx")
        Files.copy(KExcelTest::class.java.getResourceAsStream(fileName), tempFile.toPath())

        return tempFile
    }

    fun getDateTime(year: Int, month: Int, dayOfMonth: Int,
                    hour: Int, minute: Int, second: Int): Date {
        val cal = Calendar.getInstance()

        cal.set(Calendar.YEAR, year)
        cal.set(Calendar.MONTH, month - 1)
        cal.set(Calendar.DAY_OF_MONTH, dayOfMonth)
        cal.set(Calendar.HOUR_OF_DAY, hour)
        cal.set(Calendar.MINUTE, minute)
        cal.set(Calendar.SECOND, second)
        cal.set(Calendar.MILLISECOND, 0)

        return cal.time
    }

    fun getDate(year: Int, month: Int, dayOfMonth: Int): Date {
        return getDateTime(year, month, dayOfMonth, 0, 0, 0)
    }

    fun getTime(hour: Int, minute: Int, second: Int): Date {
        return getDateTime(1899, 12, 31, hour, minute, second)
    }
}
