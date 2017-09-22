package link.webarata3.kexcelapi

import org.junit.rules.TemporaryFolder

import java.io.File
import java.nio.file.Files

object TestUtil {
    @Throws(Exception::class)
    fun getTempWorkbookFile(tempFolder: TemporaryFolder, fileName: String): File {
        val tempFile = File(tempFolder.root, "temp.xlsx")
        Files.copy(KExcelTest::class.java.getResourceAsStream(fileName), tempFile.toPath())

        return tempFile
    }
}
