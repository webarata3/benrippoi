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
import org.junit.rules.ExpectedException;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;

import java.io.File;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

@RunWith(Enclosed.class)
public class BenrippoiUtilTest {
    public static class 正常系_getWorkbook {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @Test
        public void openFileNameTest() throws Exception {
            File file  = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
            Workbook wb = BenrippoiUtil.open(file.getCanonicalPath());
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
    public static class 正常系_cellIndexToCellLabelTest {
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
    public static class 異常系_cellIndexToCellLabelTest {
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(-1, 0),
            new Fixture(0, -1),
            new Fixture(-1, -1),
            new Fixture(-2, 0),
            new Fixture(0, -2),
            new Fixture(-2, -2)
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
        public void test(Fixture fixture) {
            thrown.expect(IllegalArgumentException.class);
            BenrippoiUtil.cellIndexToCellLabel(fixture.x, fixture.y);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_getCellLabelToCellIndexTest {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A1", 0, 0),
            new Fixture("B1", 1, 0),
            new Fixture("C1", 2, 0),
            new Fixture("AA1", 26, 0),
            new Fixture("AB1", 27, 0),
            new Fixture("AC1", 28, 0)
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
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            Sheet sheet = wb.getSheetAt(0);

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(fixture.toString(), cell, is(notNullValue()));
            assertThat(fixture.toString(), cell.getAddress().getColumn(), is(fixture.x));
            assertThat(fixture.toString(), cell.getAddress().getRow(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_getCellLabelToCellIndexTest {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("あ1"),
            new Fixture("AA１"),
            new Fixture("あ"),
            new Fixture("1"),
            new Fixture("A")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            Sheet sheet = wb.getSheetAt(0);

            thrown.expect(IllegalArgumentException.class);
            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
        }
    }

    @RunWith(Theories.class)
    public static class GetRowByIndex {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture(0),
            new Fixture(1),
            new Fixture(2),
            new Fixture(3),
            new Fixture(10)
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
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));
            Row row = BenrippoiUtil.getRow(sheet, fixture.y);
            assertThat(row, is(notNullValue()));
            assertThat(row.getRowNum(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_getCellByIndex {
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
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.x, fixture.y);
            assertThat(cell, is(notNullValue()));
            assertThat(cell.getAddress().getColumn(), is(fixture.x));
            assertThat(cell.getAddress().getRow(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_getCellByCellLabel {
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
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(cell.getAddress().getColumn(), is(fixture.x));
            assertThat(cell.getAddress().getRow(), is(fixture.y));
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToString {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2", "あいうえお"),
            new Fixture("C2", "123"),
            new Fixture("D2", "150.51"),
            new Fixture("F2", "true"),
            new Fixture("G2", "123150.51"),
            new Fixture("H2", ""),
            new Fixture("I2", ""),
            new Fixture("J2", "あいうえお123")
        };

        static class Fixture {
            String cellLabel;
            String expected;

            Fixture(String cellLabel, String expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToString(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_日付仮_cellToString {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E2")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(UnsupportedOperationException.class);
            BenrippoiUtil.cellToString(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToString {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("K2")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToString(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToInt {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B3", 456),
            new Fixture("C3", 123),
            new Fixture("D3", 150),
            new Fixture("G3", 369),
            new Fixture("J3", 456123)
        };

        static class Fixture {
            String cellLabel;
            int expected;

            Fixture(String cellLabel, int expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToInt(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToInt {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2"),
            new Fixture("E3"),
            new Fixture("F3"),
            new Fixture("H3"),
            new Fixture("I3"),
            new Fixture("K3")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToInt(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToDouble {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B4", 123.456),
            new Fixture("C4", 123),
            new Fixture("D4", 150.51),
            new Fixture("G4", 50.17),
            new Fixture("J4", 123123.456)
        };

        static class Fixture {
            String cellLabel;
            double expected;

            Fixture(String cellLabel, double expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }
        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToDouble(cell), is(closeTo(fixture.expected, 0.00001)));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToDouble {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2"),
            new Fixture("E4"),
            new Fixture("F4"),
            new Fixture("H4"),
            new Fixture("I4"),
            new Fixture("K4")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }
        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToDouble(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToBoolean {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("F5", true),
            new Fixture("G5", false)
        };

        static class Fixture {
            String cellLabel;
            boolean expected;

            Fixture(String cellLabel, boolean expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToBoolean(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToBoolean {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B5"),
            new Fixture("C5"),
            new Fixture("D5"),
            new Fixture("E5"),
            new Fixture("K5")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToBoolean(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToLocalDate {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E6", LocalDate.of(2015, 12, 1)),
            new Fixture("G6", LocalDate.of(2015, 12, 3))
        };

        static class Fixture {
            String cellLabel;
            LocalDate expected;

            Fixture(String cellLabel, LocalDate expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToLocalDate(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToLocalDate {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A6"),
            new Fixture("B6"),
            new Fixture("C6"),
            new Fixture("D6"),
            new Fixture("F6"),
            new Fixture("H6"),
            new Fixture("I6"),
            new Fixture("J6"),
            new Fixture("K6")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToLocalDate(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToLocalTime {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E7", LocalTime.of(10, 10, 30)),
            new Fixture("G7", LocalTime.of(12, 34, 30))
        };

        static class Fixture {
            String cellLabel;
            LocalTime expected;

            Fixture(String cellLabel, LocalTime expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToLocalTime(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToLocalTime {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A7"),
            new Fixture("B7"),
            new Fixture("C7"),
            new Fixture("D7"),
            new Fixture("F7"),
            new Fixture("H7"),
            new Fixture("I7"),
            new Fixture("J7"),
            new Fixture("K7")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToLocalTime(cell);
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_cellToLocalDateTime {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E8", LocalDateTime.of(2015, 12, 1, 10, 10, 30)),
            new Fixture("G8", LocalDateTime.of(2015,12,3, 10,10, 30))
        };

        static class Fixture {
            String cellLabel;
            LocalDateTime expected;

            Fixture(String cellLabel, LocalDateTime expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected=" + expected +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            assertThat(BenrippoiUtil.cellToLocalDateTime(cell), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_cellToLocalDateTime {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A8"),
            new Fixture("B8"),
            new Fixture("C8"),
            new Fixture("D8"),
            new Fixture("F8"),
            new Fixture("H8"),
            new Fixture("I8"),
            new Fixture("J8"),
            new Fixture("K8")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            Sheet sheet = TestUtil.getSheet(tempFolder, "book1.xlsx");
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(cell, is(notNullValue()));
            thrown.expect(PoiIllegalAccessException.class);
            BenrippoiUtil.cellToLocalDateTime(cell);
        }
    }
}
