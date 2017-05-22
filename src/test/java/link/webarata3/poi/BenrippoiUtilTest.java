package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;

import java.io.File;
import java.nio.file.Files;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.notNullValue;
import static org.junit.Assert.assertThat;

@RunWith(Enclosed.class)
public class BenrippoiUtilTest {
    public static class GetWorkbookTest {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @Test
        public void openFileNameTest() throws Exception {
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            assertThat(wb, is(notNullValue()));
            wb.close();
        }

        @Test
        public void openInputStreamTest() throws Exception {
            File file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
            Workbook wb = BenrippoiUtil.open(Files.newInputStream(file.toPath()));
            assertThat(wb, is(notNullValue()));
            wb.close();
        }
    }

    @RunWith(Theories.class)
    public static class CellIndexToCellLabelTest {
        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(0, 0, "A1"),
            new Fixture(1, 0, "B1"),
            new Fixture(2, 0, "C1"),
            new Fixture(26, 0, "AA1"),
            new Fixture(27, 0, "AB1"),
            new Fixture(28, 0, "AC1")
        };

        static class Fixture {
            int x;
            int y;
            String cellLabel;

            Fixture(int x, int y, String cellLabel) {
                this.x = x;
                this.y = y;
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "x=" + x +
                    ", y=" + y +
                    ", cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) {
            assertThat(fixture.toString(), BenrippoiUtil.cellIndexToCellLabel(fixture.x, fixture.y), is(fixture.cellLabel));
        }
    }

    @RunWith(Theories.class)
    public static class GetCellByCellLabelTest {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(0, 0, "A1"),
            new Fixture(1, 0, "B1"),
            new Fixture(2, 0, "C1"),
            new Fixture(26, 0, "AA1"),
            new Fixture(27, 0, "AB1"),
            new Fixture(28, 0, "AC1")
        };

        static class Fixture {
            int x;
            int y;
            String cellLabel;

            Fixture(int x, int y, String cellLabel) {
                this.x = x;
                this.y = y;
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "x=" + x +
                    ", y=" + y +
                    ", cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            Sheet sheet = wb.getSheetAt(0);

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(fixture.toString(), cell, is(notNullValue()));
            assertThat(fixture.toString(), cell.getAddress().getColumn(), is(fixture.x));
            assertThat(fixture.toString(), cell.getAddress().getRow(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class GetRowByIndex  {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(0),
            new Fixture(1),
            new Fixture(2),
            new Fixture(3),
            new Fixture(4)
        };

        static class Fixture {
            int y;

            Fixture(int y) {
                this.y = y;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "y=" + y + '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder);
            assertThat(sheet, is(notNullValue()));
            Row row = BenrippoiUtil.getRow(sheet, fixture.y);
            assertThat(row, is(notNullValue()));
            assertThat(row.getRowNum(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class GetCellByIndex {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(0, 0),
            new Fixture(1, 1),
            new Fixture(2, 2),
            new Fixture(3, 3),
            new Fixture(4, 4)
        };

        static class Fixture {
            int x;
            int y;

            Fixture(int x, int y) {
                this.x = x;
                this.y = y;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "x=" + x +
                    ", y=" + y +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder);
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.x, fixture.y);
            assertThat(cell, is(notNullValue()));
            assertThat(cell.getAddress().getColumn(), is(fixture.x));
            assertThat(cell.getAddress().getRow(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class GetCellByCellLabel {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A1", 0, 0),
            new Fixture("B2", 1, 1),
            new Fixture("C3", 2, 2),
            new Fixture("C4", 2, 3),
            new Fixture("C5", 2, 4)
        };

        static class Fixture {
            String cellLabel;
            int x;
            int y;

            Fixture(String cellLabel, int x, int y) {
                this.cellLabel = cellLabel;
                this.x = x;
                this.y = y;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", x=" + x +
                    ", y=" + y +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder);
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(cell.getAddress().getColumn(), is(fixture.x));
            assertThat(cell.getAddress().getRow(), is(fixture.y));
        }
    }
}
