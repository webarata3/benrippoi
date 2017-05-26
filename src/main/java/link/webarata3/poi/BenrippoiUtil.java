package link.webarata3.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

/**
 * Apache POIのラッパークラスです。
 *
 * @author webarata3
 */
public class BenrippoiUtil {
    /**
     * Excelファイルを読み込みます
     *
     * @param fileName Excelファイル名。拡張子で読み込むフォーマットが決まります
     * @return Excel Workbook
     * @throws IOException            ファイルがない場合等
     * @throws InvalidFormatException フォーマットの例外
     */
    public static Workbook open(String fileName) throws IOException, InvalidFormatException {
        InputStream is = Files.newInputStream(Paths.get(fileName));
        return open(is);
    }

    /**
     * Excelファイルを読み込みます
     *
     * @param is ExcelファイルのInputStream
     * @return Excel Workbook
     * @throws IOException            ファイルがない場合等
     * @throws InvalidFormatException フォーマットの例外
     */
    public static Workbook open(InputStream is) throws IOException, InvalidFormatException {
        return WorkbookFactory.create(is);
    }

    /**
     * Excelのセルのインデックスをセルのラベル（A1、B2）に変更します。
     *
     * @param x 列番号（0〜）
     * @param y 行番号（0〜）
     * @return セルのラベル
     * @throws IllegalArgumentException x、yのいずれかが0未満の場合
     */
    public static String cellIndexToCellLabel(int x, int y)  {
        if (x < 0) throw new IllegalArgumentException("xは0以上でなければなりません: " + x);
        if (y < 0) throw new IllegalArgumentException("yは0以上でなければなりません: " + y);

        String cellName = dec26(x, 0);
        return cellName + (y + 1);
    }

    private static String dec26(int num, int first) {
        return (num > 25 ? dec26(num / 26, 1) : "") + String.valueOf((char) ('A' + (num - first) % 26));
    }

    /**
     * Row（行）の取得
     *
     * @param sheet シート
     * @param y     行番号（0〜）
     * @return Row（行）
     */
    public static Row getRow(Sheet sheet, int y) {
        Row row = sheet.getRow(y);
        if (row != null) {
            return row;
        }
        return sheet.createRow(y);
    }

    /**
     * Cellの取得
     *
     * @param sheet シート
     * @param x     列番号（0〜）
     * @param y     行番号（0〜）
     * @return Cell
     */
    public static Cell getCell(Sheet sheet, int x, int y) {
        Row row = getRow(sheet, y);
        Cell cell = row.getCell(x);
        if (cell != null) {
            return cell;
        }
        return row.createCell(x, CellType.BLANK);
    }

    /**
     * セルのラベル（A1、B2）のセルの取得
     *
     * @param sheet     シート
     * @param cellLabel セルのラベル（A1、B2）
     * @return Cell
     */
    public static Cell getCell(Sheet sheet, String cellLabel) {
        Pattern p1 = Pattern.compile("([a-zA-Z]+)([0-9]+)");
        Matcher matcher = p1.matcher(cellLabel);
        if (!matcher.find()) throw new IllegalArgumentException("セルラベルに「" + cellLabel + "」は指定できません。");

        // 上の位から計算するため、Cell LabelのAB1のABの部分を逆にする。
        String reverseString = new StringBuilder(matcher.group(1).toUpperCase()).reverse().toString();
        // Aを1～Zを26として、ラベルを数値に変換する。
        // 26進数なので、上位の桁は26倍する
        int x = IntStream.range(0, reverseString.length()).map((i) -> {
            int delta = reverseString.charAt(i) - 'A' + 1;
            return delta * (int) Math.pow(26.0, (double) i);
        }).reduce(-1, (v1, v2) -> v1 + v2); // 集計するが、0始まりなので、合計を-1する

        return getCell(sheet, x, Integer.parseInt(matcher.group(2)) - 1);
    }

    /**
     * セルの値をString型で取得する
     *
     * @param cell セル
     * @return String型の値
     */
    public static String cellToString(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toStr();
    }

    /**
     * セルの値をint型で取得する。小数は切り捨てられる
     *
     * @param cell セル
     * @return int型の値
     */
    public static int cellToInt(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toInt();
    }

    /**
     * セルの値をdouble型で取得する。
     *
     * @param cell セル
     * @return double型の値
     */
    public static double cellToDouble(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toDouble();
    }

    /**
     * セルの値をboolean型で取得する。
     *
     * @param cell セル
     * @return boolean型の値
     */
    public static boolean cellToBoolean(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toBoolean();
    }

    /**
     * セルの値をLocalDate型で取得する。
     *
     * @param cell セル
     * @return LocalDate型の値
     */
    public static LocalDate cellToLocalDate(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toLocalDate();
    }

    /**
     * セルの値をLocalTime型で取得する。
     *
     * @param cell セル
     * @return LocalTime型の値
     */
    public static LocalTime cellToLocalTime(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toLocalTime();
    }

    /**
     * セルの値をLocalDateTime型で取得する。
     *
     * @param cell セル
     * @return LocalDateTime型の値
     */
    public static LocalDateTime cellToLocalDateTime(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.toLocalDateTime();
    }
}
