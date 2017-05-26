package link.webarata3.poi;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

public class BenriWorkbookFactoryTest {
    @Rule
    public TemporaryFolder tempFolder = new TemporaryFolder();

    @Test
    public void 正常系_create_inputStream() throws Exception {
        File file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
        BenriWorkbook bwb = BenriWorkbookFactory.create(Files.newInputStream(file.toPath()));
        assertThat(bwb, is(notNullValue()));
    }

    @Test
    public void 正常系_create_fileName() throws Exception {
        File file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx");
        BenriWorkbook bwb = BenriWorkbookFactory.create(file.getCanonicalPath());
        assertThat(bwb, is(notNullValue()));
    }

    @Test
    public void 正常系_createBlank() throws Exception {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        assertThat(bwb, is(notNullValue()));
    }
}
