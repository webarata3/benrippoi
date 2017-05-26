package link.webarata3.poi;


import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.rules.TemporaryFolder;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

public class BenriSheetTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    public void 正常系_sheet() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            BenriSheet sheet = new BenriSheet(wb.getSheet("Sheet1"));
            assertThat(sheet, is(notNullValue()));
        }
    }

    public void 正常系_cellIndex() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            BenriSheet sheet = new BenriSheet(wb.getSheet("Sheet1"));
            assertThat(sheet, is(notNullValue()));
            BenriCell cell = sheet.cell(1, 1);
            assertThat(cell, is(notNullValue()));
        }
    }

    public void 正常系_cellLabel() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            BenriSheet sheet = new BenriSheet(wb.getSheet("Sheet1"));
            assertThat(sheet, is(notNullValue()));
            BenriCell cell = sheet.cell("A1");
            assertThat(cell, is(notNullValue()));
        }
    }

}
