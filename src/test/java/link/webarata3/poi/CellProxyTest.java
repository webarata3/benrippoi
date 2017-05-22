package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Rule;
import org.junit.experimental.runners.Enclosed;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.rules.ExpectedException;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.notNullValue;
import static org.junit.Assert.assertThat;

@RunWith(Enclosed.class)
public class CellProxyTest {
    @RunWith(Theories.class)
    public static class 正常系_toStr {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2", "あいうえお"),
            new Fixture("C2", "123"),
            new Fixture("D2", "150.51"),
            new Fixture("E2", "42339"),
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
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            assertThat(wb, is(notNullValue()));

            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(fixture.toString(), cell, is(notNullValue()));

            CellProxy cellProxy = new CellProxy(cell);
            assertThat(cellProxy.toStr(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_toStr {
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
            Workbook wb = TestUtil.getTempWorkbook(tempFolder, "book1.xlsx");
            assertThat(wb, is(notNullValue()));

            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet, is(notNullValue()));

            Cell cell = BenrippoiUtil.getCell(sheet, fixture.cellLabel);
            assertThat(fixture.toString(), cell, is(notNullValue()));

            CellProxy cellProxy = new CellProxy(cell);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toStr();
        }
    }
}
