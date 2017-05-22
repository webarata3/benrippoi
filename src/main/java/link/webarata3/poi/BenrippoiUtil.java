package link.webarata3.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
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
     */
    public static String cellIndexToCellLabel(int x, int y) {
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
        Row row = sheet.getRow(y);
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
        matcher.find();

        String reverseString = new StringBuilder(matcher.group(1).toUpperCase()).reverse().toString();
        int x = IntStream.range(0, reverseString.length()).map((i) -> {
            int delta = reverseString.charAt(i) - 'A' + 1;
            return delta * (int) Math.pow(26.0, (double) i);
        }).reduce(-1, (v1, v2) -> v1 + v2);

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
        return cellProxy.cellToString();
    }

    /**
     * セルの値をint型で取得する。小数は切り捨てられる
     *
     * @param cell セル
     * @return int型の値
     */
    public static int cellToInt(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.cellToInt();
    }

    /**
     * セルの値をdouble型で取得する。
     *
     * @param cell セル
     * @return double型の値
     */
    public static double cellToDouble(Cell cell) {
        CellProxy cellProxy = new CellProxy(cell);
        return cellProxy.cellToDouble();
    }
}
