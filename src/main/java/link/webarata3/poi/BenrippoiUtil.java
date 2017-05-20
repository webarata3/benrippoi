package link.webarata3.poi;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

/**
 * Apache POIのラッパークラスです。
 *
 * @author webarata3
 */
public class BenrippoiUtil {
    public static Optional<Workbook> open(String fileName) {
        try {
            InputStream is = Files.newInputStream(Paths.get(fileName));
            return open(is);
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    public static Optional<Workbook> open(InputStream is) {
        Objects.requireNonNull(is, "InputStreamにnullは許可されていません");
        try (Workbook wb = WorkbookFactory.create(is)) {
            return Optional.of(wb);
        } catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
            return Optional.empty();
        }
    }

    public static String cellIndexToCellLabel(int x, int y) {
        String cellName = dec26(x, 0);
        return cellName + (y + 1);
    }

    private static String dec26(int num, int first) {
        return (num > 25 ? dec26(num / 26, 1) : "") + String.valueOf((char) ('A' + (num - first) % 26));
    }

    public static Cell getCell(Sheet sheet, String cellLabel) {
        Pattern p1 = Pattern.compile("([a-zA-Z]+)([0-9]+)");
        Matcher matcher = p1.matcher(cellLabel);
        matcher = null;
        matcher.find();

        String reverseString = new StringBuilder(matcher.group(1).toUpperCase()).reverse().toString();
        int x = IntStream.range(0, reverseString.length()).map((i) -> {
            int delta = reverseString.charAt(i) - 'A' + 1;
            return delta * (int) Math.pow(26.0, (double) i);
        }).reduce(-1, (v1, v2) -> v1 + v2);

        return getCell(sheet, x, Integer.parseInt(matcher.group(2)) - 1);
    }

    public static Row getRow(Sheet sheet, int n) {
        Row row = sheet.getRow(n);
        if (row != null) {
            return row;
        }
        return sheet.createRow(n);
    }

    public static Cell getCell(Sheet sheet, int x, int y) {
        Row row = sheet.getRow(y);
        Cell cell = row.getCell(x);
        if (cell != null) {
            return cell;
        }
        return row.createCell(x, CellType.BLANK);
    }

    public static String normalizeNumericString(double numeric) {
        // 44.0のような数値を44として取得するために、入力された数値と小数点以下を切り捨てた数値が
        // 一致した場合には、intにキャストして、小数点以下が表示されないようにしている
        if (numeric == Math.ceil(numeric)) {
            return String.valueOf((int) numeric);
        }
        return String.valueOf(numeric);
    }

    public static CellValue getFomulaCellValue(Cell cell) {
        Workbook wb = cell.getSheet().getWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        FormulaEvaluator evaluator = helper.createFormulaEvaluator();
        return evaluator.evaluate(cell);
    }

    public static String cellToString(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return normalizeNumericString(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            case FORMULA:
                CellValue cellValue = getFomulaCellValue(cell);
                switch (cellValue.getCellTypeEnum()) {
                    case STRING:
                        return cellValue.getStringValue();
                    case NUMERIC:
                        return normalizeNumericString(cellValue.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cellValue.getBooleanValue());
                    case BLANK:
                        return "";
                    default: // _NONE, ERROR
                        throw new PoiIllegalAccessException("cellはStringに変換できません");
                }
            default: // _NONE, ERROR
                throw new PoiIllegalAccessException("cellはStringに変換できません");
        }
    }

    private static int stringToInt(String value) {
        try {
            return (int) Double.parseDouble(value);
        } catch (NumberFormatException e) {
            throw new IllegalStateException("cellはintに変換できません");
        }
    }

    public static int cellToInt(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return stringToInt(cell.getStringCellValue());
            case NUMERIC:
                return (int) cell.getNumericCellValue();
            case FORMULA:
                CellValue cellValue = getFomulaCellValue(cell);
                switch (cellValue.getCellTypeEnum()) {
                    case STRING:
                        return stringToInt(cellValue.getStringValue());
                    case NUMERIC:
                        return (int) cellValue.getNumberValue();
                    default:
                        throw new PoiIllegalAccessException("cellはintに変換できません");
                }
            default:
                throw new PoiIllegalAccessException("cellはintに変換できません");
        }
    }

    private static double stringToDouble(String value) {
        try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            throw new PoiIllegalAccessException("cellはdoubleに変換できません");
        }
    }

    public static double cellToDouble(Cell cell) {
        switch(cell.getCellTypeEnum()) {
            case STRING:
                return stringToDouble(cell.getStringCellValue());
            case NUMERIC:
                return cell.getNumericCellValue();
            case FORMULA:
                CellValue cellValue = getFomulaCellValue(cell);
                switch(cellValue.getCellTypeEnum()) {
                    case STRING:
                        return stringToDouble(cell.getStringCellValue());
                    case NUMERIC:
                        return cell.getNumericCellValue();
                    default:
                        throw new PoiIllegalAccessException("cellはdoubleに変換できません");
                }
            default:
                throw new PoiIllegalAccessException("cellはdoubleに変換できません");
        }
    }
}
