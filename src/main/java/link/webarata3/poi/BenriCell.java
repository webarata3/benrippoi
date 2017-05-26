package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;

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

    public void set(String value) {
        cell.setCellValue(value);
    }

    public void set(int value) {
        cell.setCellValue(value);
    }

    public void set(double value) {
        cell.setCellValue(value);
    }

    public void set(boolean value) {
        cell.setCellValue(value);
    }

    private Date localDataTimeToDate(LocalDateTime localDateTime) {
        ZoneId zone = ZoneId.systemDefault();
        ZonedDateTime zonedDateTime = ZonedDateTime.of(localDateTime, zone);

        Instant instant = zonedDateTime.toInstant();
        Date date = Date.from(instant);

        return date;
    }

    public void set(LocalDate value) {
        cell.setCellValue(localDataTimeToDate(value.atStartOfDay()));
    }

    public void set(LocalTime value) {
        cell.setCellValue(localDataTimeToDate(value.atDate(LocalDate.of(1900, 1,0))));
    }
}
