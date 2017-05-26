package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

public class BenriCellTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    @Test
    public void 正常系_toStr() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            String actual = wbb.sheet("Sheet1").cell("B2").toStr();
            assertThat(actual, is("あいうえお"));
        }
    }

    @Test
    public void 正常系_toInt() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            int actual = wbb.sheet("Sheet1").cell("C3").toInt();
            assertThat(actual, is(123));
        }
    }

    @Test
    public void 正常系_toDouble() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            double actual = wbb.sheet("Sheet1").cell("D4").toDouble();
            assertThat(actual, is(closeTo(150.51, 0.0000001)));
        }
    }

    @Test
    public void 正常系_toBoolean() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            boolean actual = wbb.sheet("Sheet1").cell("F5").toBoolean();
            assertThat(actual, is(true));
        }
    }

    @Test
    public void 正常系_toLocalDate() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            LocalDate actual = wbb.sheet("Sheet1").cell("E6").toLocalDate();
            assertThat(actual, is(LocalDate.of(2015, 12, 1)));
        }
    }

    @Test
    public void 正常系_toLocalTime() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            LocalTime actual = wbb.sheet("Sheet1").cell("E7").toLocalTime();
            assertThat(actual, is(LocalTime.of(10, 10, 30)));
        }
    }

    @Test
    public void 正常系_toLocalDateTime() throws Exception {
        Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
        try (BenriWorkbook wbb = new BenriWorkbook(wb)) {
            LocalDateTime actual = wbb.sheet("Sheet1").cell("E8").toLocalDateTime();
            assertThat(actual, is(LocalDateTime.of(2015, 12, 1, 10, 10, 30)));
        }
    }

    @Test
    public void 正常系_set_String() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set("あいうえお");
        assertThat(sheet.cell("A1").toStr(), is("あいうえお"));
    }

    @Test
    public void 正常系_set_int() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(12345);
        assertThat(sheet.cell("A1").toInt(), is(12345));
    }

    @Test
    public void 正常系_set_double() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(12345.51);
        assertThat(sheet.cell("A1").toDouble(), is(closeTo(12345.51, 0.000001)));
    }

    @Test
    public void 正常系_set_boolean() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(true);
        assertThat(sheet.cell("A1").toBoolean(), is(true));
    }

    @Test
    public void 正常系_set_LocalDate() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(LocalDate.of(2017, 5, 27));
        assertThat(sheet.cell("A1").toLocalDate(), is(LocalDate.of(2017, 5, 27)));
    }

    @Test
    public void 正常系_set_LocalTime() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(LocalTime.of(14, 10, 59));
        assertThat(sheet.cell("A1").toLocalTime(), is(LocalTime.of(14, 10, 59)));
    }

    @Test
    public void 正常系_set_LocalDateTime() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
        BenriSheet sheet = bwb.createSheet("テストシート");
        assertThat(sheet, is(notNullValue()));
        sheet.cell("A1").set(LocalDateTime.of(2017, 5, 27, 14, 10, 59));
        assertThat(sheet.cell("A1").toLocalDateTime(), is(LocalDateTime.of(2017, 5, 27, 14, 10, 59)));
    }
}
