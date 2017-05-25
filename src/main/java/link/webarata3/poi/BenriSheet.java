package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Sheet;

public class BenriSheet {
    private Sheet sheet;

    public BenriSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public static BenriSheet create(Sheet sheet) {
        return new BenriSheet(sheet);
    }

    public BenriCell cell(int x, int y) {
        return new BenriCell(sheet, x, y);
    }
    public BenriCell cell(String cellLabel) {
        return new BenriCell(sheet, cellLabel);
    }
}
