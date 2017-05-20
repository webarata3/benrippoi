package link.webarata3.poi;

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
import java.io.FileInputStream;
import java.nio.file.Files;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

@RunWith(Enclosed.class)
public class BenrippoiUtilTest {
    private static File getTempWorkbookFile(TemporaryFolder tempFolder, String fileName) throws Exception {
        File tempFile = new File(tempFolder.getRoot(), "temp.xlsx");
        Files.copy(BenrippoiUtil.class.getResourceAsStream(fileName), tempFile.toPath());

        return tempFile;
    }

    private static Workbook getTempWorkbook(TemporaryFolder tempFolder, String fileName) throws Exception {
        File tempFile = getTempWorkbookFile(tempFolder, fileName);
        return BenrippoiUtil.open(Files.newInputStream(tempFile.toPath()));
    }

    public static class GetWorkbookTest {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @Test
        public void openFileNameTest() throws Exception {
            Workbook wb = BenrippoiUtilTest.getTempWorkbook(tempFolder, "book1.xlsx");
            wb.close();
        }

        public void openInputStreamTest() throws Exception {
            File file = getTempWorkbookFile(tempFolder, "book1.xlsx");
            Workbook sb = BenrippoiUtil.open(Files.newInputStream(file.toPath()));
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
}
