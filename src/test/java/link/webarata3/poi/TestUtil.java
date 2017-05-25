package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

public class TestUtil {
    public static File getTempWorkbookFile(TemporaryFolder tempFolder, String fileName) throws Exception {
        File tempFile = new File(tempFolder.getRoot(), "temp.xlsx");
        Files.copy(BenrippoiUtil.class.getResourceAsStream(fileName), tempFile.toPath());

        return tempFile;
    }

    public static Workbook getTempWorkbook(TemporaryFolder tempFolder, String fileName) throws Exception {
        File tempFile = getTempWorkbookFile(tempFolder, fileName);
        return BenrippoiUtil.open(Files.newInputStream(tempFile.toPath()));
    }

    public static Sheet getSheet(TemporaryFolder tempFolder, String fileName) throws Exception {
        Workbook wb = getTempWorkbook(tempFolder, fileName);
        return wb.getSheetAt(0);
    }

    public static CellProxy getCellProxy(TemporaryFolder tempFolder, String fileName, String cellLabel) throws Exception {
        Sheet sheet = getSheet(tempFolder, fileName);
        Cell cell = BenrippoiUtil.getCell(sheet, cellLabel);
        return new CellProxy(cell);
    }
}
