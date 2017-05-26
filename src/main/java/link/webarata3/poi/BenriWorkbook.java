package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

public class BenriWorkbook implements Closeable {
    private Workbook wb;

    /**
     * シートのキャッシュ
     */
    private Map<String, BenriSheet> benriSheetMap;

    public BenriWorkbook(Workbook wb) {
        this.wb = wb;
        benriSheetMap = new HashMap<>();
    }

    @Override
    public void close() throws IOException {
        wb.close();
    }

    public void write(String fileName) throws IOException {
        wb.write(Files.newOutputStream(Paths.get(fileName)));
    }

    public void write(OutputStream os) throws IOException {
        wb.write(os);
    }

    public BenriSheet sheet(String sheetName) {
        return benriSheetMap.computeIfAbsent(sheetName, k -> new BenriSheet(wb.getSheet(k)));
    }

    public BenriSheet sheetAt(int index) {
        String sheetName = wb.getSheetAt(index).getSheetName();
        return sheet(sheetName);
    }

    public BenriSheet createSheet(String sheetName) {
        Sheet sheet = wb.createSheet(sheetName);
        BenriSheet benriSheet = new BenriSheet(sheet);
        benriSheetMap.put(sheetName, benriSheet);
        return benriSheet;
    }
}
