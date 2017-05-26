package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

public class BenriCell {
    private Cell cell;

    public BenriCell(Sheet sheet, int x, int y) {
        cell = BenrippoiUtil.getCell(sheet, x, y);
    }

    public BenriCell(Sheet sheet, String cellLabel) {
        cell = BenrippoiUtil.getCell(sheet, cellLabel);
    }

    public String toStr() {
        return BenrippoiUtil.cellToString(cell);
    }

    public int toInt() {
        return BenrippoiUtil.cellToInt(cell);
    }

    public double toDouble() {
        return BenrippoiUtil.cellToDouble(cell);
    }

    public boolean toBoolean() {
        return BenrippoiUtil.cellToBoolean(cell);
    }

    public LocalDate toLocalDate() {
        return BenrippoiUtil.cellToLocalDate(cell);
    }

    public LocalTime toLocalTime() {
        return BenrippoiUtil.cellToLocalTime(cell);
    }

    public LocalDateTime toLocalDateTime() {
        return BenrippoiUtil.cellToLocalDateTime(cell);
    }
}
