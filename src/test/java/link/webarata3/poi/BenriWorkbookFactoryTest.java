package link.webarata3.poi;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

public class BenriWorkbookFactoryTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    @Test
    public void 正常系_create_inputStream() throws Exception {
        File file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
        BenriWorkbook bwb = BenriWorkbookFactory.create(Files.newInputStream(file.toPath()));
    }

    @Test
    public void 正常系_create_fileName() throws Exception {
        File file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
        BenriWorkbook bwb = BenriWorkbookFactory.create(file.getCanonicalPath());
    }
}
