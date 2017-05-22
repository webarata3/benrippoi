package link.webarata3.poi;

import org.apache.poi.ss.usermodel.*;

public class CellProxy {
    private Cell cell;
    private CellValue cellValue;

    public CellProxy(Cell cell) {
        if (cell.getCellTypeEnum() == CellType.FORMULA) {
            this.cellValue = getFomulaCellValue(cell);
        } else {
            this.cell = cell;
        }
    }

    private CellType getCellTypeEnum() {
        if (cell == null) {
            return cellValue.getCellTypeEnum();
        } else {
            return cell.getCellTypeEnum();
        }
    }

    private String getStringCellValue() {
        if (cell == null) {
            return cellValue.getStringValue();
        } else {
            return cell.getStringCellValue();
        }
    }

    private double getNumericCellValue() {
        if (cell == null) {
            return cellValue.getNumberValue();
        } else {
            return cell.getNumericCellValue();
        }
    }

    private boolean getBooleanCellValue() {
        if (cell == null) {
            return cellValue.getBooleanValue();
        } else {
            return cell.getBooleanCellValue();
        }
    }

    /**
     * 数値の正規化
     *
     * @param numeric 正規化する数値
     * @return 正規化した数値
     */
    private String normalizeNumericString(double numeric) {
        // 44.0のような数値を44として取得するために、入力された数値と小数点以下を切り捨てた数値が
        // 一致した場合には、intにキャストして、小数点以下が表示されないようにしている
        if (numeric == Math.ceil(numeric)) {
            return String.valueOf((int) numeric);
        }
        return String.valueOf(numeric);
    }

    private int stringToInt(String value) {
        try {
            return (int) Double.parseDouble(value);
        } catch (NumberFormatException e) {
            throw new IllegalStateException("cellはintに変換できません");
        }
    }

    private double stringToDouble(String value) {
        try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            throw new PoiIllegalAccessException("cellはdoubleに変換できません");
        }
    }

    /**
     * 計算式のセルの値の取得
     *
     * @param cell 計算式があるセル
     * @return CellValue
     */
    private CellValue getFomulaCellValue(Cell cell) {
        Workbook wb = cell.getSheet().getWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        FormulaEvaluator evaluator = helper.createFormulaEvaluator();
        return evaluator.evaluate(cell);
    }

    public String toStr() {
        switch (getCellTypeEnum()) {
            case STRING:
                return getStringCellValue();
            case NUMERIC:
                return normalizeNumericString(getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(getBooleanCellValue());
            case BLANK:
                return "";
            default: // _NONE, ERROR
                throw new PoiIllegalAccessException("cellはStringに変換できません");
        }
    }

    public int toInt() {
        switch (getCellTypeEnum()) {
            case STRING:
                return stringToInt(getStringCellValue());
            case NUMERIC:
                return (int) getNumericCellValue();
            default:
                throw new PoiIllegalAccessException("cellはintに変換できません");
        }
    }

    public double toDouble() {
        switch (getCellTypeEnum()) {
            case STRING:
                return stringToDouble(getStringCellValue());
            case NUMERIC:
                return getNumericCellValue();
            default:
                throw new PoiIllegalAccessException("cellはdoubleに変換できません");
        }
    }

    public boolean toBoolean() {
        switch (getCellTypeEnum()) {
            case BOOLEAN:
                return getBooleanCellValue();
            default:
                throw new PoiIllegalAccessException("cellはdoubleに変換できません");
        }
    }
}
