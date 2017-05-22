package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

public class WrapperCell {
    private Cell cell;

    public WrapperCell(Sheet sheet, int x, int y) {
        cell = BenrippoiUtil.getCell(sheet, x, y);
    }

    public WrapperCell(Sheet sheet, String cellLabel) {
        cell = BenrippoiUtil.getCell(sheet, cellLabel);
    }

    public int toInt() {
       return BenrippoiUtil.cellToInt(cell);
    }
}
