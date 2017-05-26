package link.webarata3.poi;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

public class BenriWorkbookFactory {
    public static BenriWorkbook create(String fileName) throws IOException, InvalidFormatException {
        InputStream is = Files.newInputStream(Paths.get(fileName));
        return create(is);
    }

    public static BenriWorkbook create(InputStream is) throws IOException, InvalidFormatException {
        Workbook wb = BenrippoiUtil.open(is);
        return new BenriWorkbook(wb);
    }

    public static BenriWorkbook createBlank() {
        Workbook wb = new XSSFWorkbook();
        return new BenriWorkbook(wb);
    }
}
