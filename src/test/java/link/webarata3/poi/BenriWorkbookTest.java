package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.nio.file.Files;
import java.nio.file.Paths;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

public class BenriWorkbookTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    @Test
    public void 正常系_BenriWorkbook() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        BenriWorkbook bwb = new BenriWorkbook(wb);
        assertThat(bwb, is(notNullValue()));
        bwb.close();
    }

    @Test
    public void 正常系_sheet() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook bwb = new BenriWorkbook(wb)) {
            BenriSheet sheet = bwb.sheet("Sheet1");
            assertThat(sheet, is(notNullValue()));
        }
    }

    @Test
    public void 正常系_sheetAt() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook bwb = new BenriWorkbook(wb)) {
            BenriSheet sheet = bwb.sheetAt(0);
            assertThat(sheet, is(notNullValue()));
        }
    }

    @Test
    public void 正常系_createSheet() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook bwb = new BenriWorkbook(wb)) {
            BenriSheet sheet = bwb.createSheet("あいうえお");
            assertThat(sheet, is(notNullValue()));
        }
    }

    @Test
    public void 正常系_save_fileName() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        bwb.write(Paths.get(tempFolder.getRoot().getCanonicalPath(), "test.xlsx").toFile().getCanonicalPath());
    }

    @Test
    public void 正常系_save_outputStream() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        bwb.write(Files.newOutputStream(Paths.get(tempFolder.getRoot().getCanonicalPath(), "test.xlsx")));
    }
}
