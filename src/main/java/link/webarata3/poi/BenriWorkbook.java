package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class BenriWorkbook {
    private Workbook wb;

    /**
     * シートのキャッシュ
     */
    private Map<String, BenriSheet> benriSheetMap;

    public BenriWorkbook(Workbook wb) {
        this.wb = wb;
        benriSheetMap = new HashMap<String, BenriSheet>();
    }

    public BenriSheet sheet(String sheetName) {
        return benriSheetMap.computeIfAbsent(sheetName, k -> new BenriSheet(wb.getSheet(k)));
    }

    public BenriSheet sheetAt(int index) {
        String sheetName = wb.getSheetAt(index).getSheetName();
        return sheet(sheetName);
    }
}
