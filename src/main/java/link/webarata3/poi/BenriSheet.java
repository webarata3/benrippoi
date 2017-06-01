package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Sheetを便利に扱うためのクラス
 */
public class BenriSheet {
    private Sheet sheet;

    /**
     * コンストラクタ
     *
     * @param sheet シート
     */
    public BenriSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * シートから、指定の行列のBenriCellを取得する
     *
     * @param x 列番号（0〜）
     * @param y 行番号（0〜）
     * @return BenriCell
     */
    public BenriCell cell(int x, int y) {
        return new BenriCell(sheet, x, y);
    }

    /**
     * シートから、指定のラベル（A1、B3）のBenriCellを取得する
     *
     * @param cellLabel セルのラベル
     * @return BenriCell
     */
    public BenriCell cell(String cellLabel) {
        return new BenriCell(sheet, cellLabel);
    }
}
