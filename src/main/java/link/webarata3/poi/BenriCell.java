package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;

/**
 * Cellを便利に扱うクラス
 */
public class BenriCell {
    private Cell cell;

    /**
     * シートから(x, y)のセルを作成する
     *
     * @param sheet 対象シート
     * @param x     列番号（0〜）
     * @param y     行番号（0〜）
     */
    public BenriCell(Sheet sheet, int x, int y) {
        cell = BenrippoiUtil.getCell(sheet, x, y);
    }

    /**
     * シートからセルラベル（A1、B2）のセルを作成する
     *
     * @param sheet
     * @param cellLabel
     */
    public BenriCell(Sheet sheet, String cellLabel) {
        cell = BenrippoiUtil.getCell(sheet, cellLabel);
    }

    /**
     * 現在のセルのString型で取得する
     *
     * @return String型の値
     */
    public String toStr() {
        return BenrippoiUtil.cellToString(cell);
    }

    /**
     * セルの値をint型で取得する。小数は切り捨てられる
     *
     * @return int型の値
     */
    public int toInt() {
        return BenrippoiUtil.cellToInt(cell);
    }

    /**
     * セルの値をdouble型で取得する
     *
     * @return double型の値
     */
    public double toDouble() {
        return BenrippoiUtil.cellToDouble(cell);
    }

    /**
     * セルの値をboolean型で取得する
     *
     * @return boolean型の値
     */
    public boolean toBoolean() {
        return BenrippoiUtil.cellToBoolean(cell);
    }

    /**
     * セルの値をLocalDate型で取得する
     *
     * @return LocalDate型の値
     */
    public LocalDate toLocalDate() {
        return BenrippoiUtil.cellToLocalDate(cell);
    }

    /**
     * セルの値をLocalTime型で取得する
     *
     * @return LocalTime型の値
     */
    public LocalTime toLocalTime() {
        return BenrippoiUtil.cellToLocalTime(cell);
    }

    /**
     * セルの値をLocalDateTime型で取得する
     *
     * @return LocalDateTime型の値
     */
    public LocalDateTime toLocalDateTime() {
        return BenrippoiUtil.cellToLocalDateTime(cell);
    }

    /**
     * String型の値をセットする
     *
     * @param value String型の値
     */
    public void set(String value) {
        cell.setCellValue(value);
    }

    /**
     * int型の値をセットする
     *
     * @param value int型の値
     */
    public void set(int value) {
        cell.setCellValue(value);
    }

    /**
     * double型の値をセットする
     *
     * @param value double型の値
     */
    public void set(double value) {
        cell.setCellValue(value);
    }

    /**
     * boolean型の値をセットする
     *
     * @param value boolean型の値
     */
    public void set(boolean value) {
        cell.setCellValue(value);
    }

    private Date localDataTimeToDate(LocalDateTime localDateTime) {
        ZoneId zone = ZoneId.systemDefault();
        ZonedDateTime zonedDateTime = ZonedDateTime.of(localDateTime, zone);

        Instant instant = zonedDateTime.toInstant();
        return Date.from(instant);
    }

    private void setDateFormat(String format) {
        Workbook wb = cell.getSheet().getWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
        cell.setCellStyle(cellStyle);
    }

    /**
     * LocalDate型の値をセットする
     *
     * @param value LocalDateTime型の値
     */
    public void set(LocalDate value) {
        setDateFormat("yyyy/mm/dd");
        cell.setCellValue(localDataTimeToDate(value.atStartOfDay()));
    }

    /**
     * LocalTime型の値をセットする
     *
     * @param value LocalTime型の値
     */
    public void set(LocalTime value) {
        setDateFormat("hh:mm:ss");
        cell.setCellValue(localDataTimeToDate(value.atDate(LocalDate.of(1900, 1, 1))));
    }

    /**
     * LocalDateTime型の値をセットする
     *
     * @param value LocalDateTime型の値
     */
    public void set(LocalDateTime value) {
        setDateFormat("yyyy/mm/dd hh:mm:ss");
        cell.setCellValue(localDataTimeToDate(value));
    }
}
