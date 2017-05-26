package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

public class BenriWorkbookTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    @Test
    public void 正常系_BenriWorkbook() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        BenriWorkbook wbb = new BenriWorkbook(wb);
        assertThat(wbb, is(notNullValue()));
        wbb.close();
    }

    @Test
    public void 正常系_sheet() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            BenriSheet sheet = wbb.sheet("Sheet1");
            assertThat(sheet, is(notNullValue()));
        }
    }

    @Test
    public void 正常系_sheetAt() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            BenriSheet sheet = wbb.sheetAt(0);
            assertThat(sheet, is(notNullValue()));
        }
    }
}
