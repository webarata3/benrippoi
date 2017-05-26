package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
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

    public BenriSheet sheet(String sheetName) {
        return benriSheetMap.computeIfAbsent(sheetName, k -> new BenriSheet(wb.getSheet(k)));
    }

    public BenriSheet sheetAt(int index) {
        String sheetName = wb.getSheetAt(index).getSheetName();
        return sheet(sheetName);
    }
}
