package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Sheet;

public class WrapperSheet {
    private Sheet sheet;

    public WrapperSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public static WrapperSheet sheet(Sheet sheet) {
        return new WrapperSheet(sheet);
    }

    public WrapperCell cell(int x, int y) {
        return new WrapperCell(sheet, x, y);
    }
    public WrapperCell cell(String cellLabel) {
        return new WrapperCell(sheet, cellLabel);
    }
}
