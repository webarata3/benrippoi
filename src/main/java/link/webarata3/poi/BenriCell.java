package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

public class BenriCell {
    private Cell cell;

    public BenriCell(Sheet sheet, int x, int y) {
        cell = BenrippoiUtil.getCell(sheet, x, y);
    }

    public BenriCell(Sheet sheet, String cellLabel) {
        cell = BenrippoiUtil.getCell(sheet, cellLabel);
    }

    public int toInt() {
        return BenrippoiUtil.cellToInt(cell);
    }

    public String toStr() {
        return BenrippoiUtil.cellToString(cell);
    }
}
