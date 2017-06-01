package link.webarata3.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * 便利なWorkbookの昨日を提供する
 */
public class BenriWorkbook implements Closeable {
    private Workbook wb;

    /**
     * シートのキャッシュ
     */
    private Map<String, BenriSheet> benriSheetMap;

    /**
     * コンストラクタ
     *
     * @param wb ワークブック
     */
    public BenriWorkbook(Workbook wb) {
        this.wb = wb;
        benriSheetMap = new HashMap<>();
    }

    /**
     * ワークブックのclose
     *
     * @throws IOException
     */
    @Override
    public void close() throws IOException {
        wb.close();
    }

    /**
     * ワークブックのファイルへの書き出し
     *
     * @param fileName ファイル名
     * @throws IOException
     */
    public void write(String fileName) throws IOException {
        wb.write(Files.newOutputStream(Paths.get(fileName)));
    }

    /**
     * ワークブックへのファイルへの書き出し
     *
     * @param os
     * @throws IOException
     */
    public void write(OutputStream os) throws IOException {
        wb.write(os);
    }

    /**
     * ワークブックからシート名を指定してシートを取り出す
     *
     * @param sheetName シート名
     * @return シート
     */
    public BenriSheet sheet(String sheetName) {
        return benriSheetMap.computeIfAbsent(sheetName, k -> new BenriSheet(wb.getSheet(k)));
    }

    /**
     * ワークブックからシート番号を指定してシートを取り出す
     *
     * @param index シート番号
     * @return シート
     */
    public BenriSheet sheetAt(int index) {
        String sheetName = wb.getSheetAt(index).getSheetName();
        return sheet(sheetName);
    }

    /**
     * ワークブックに新しいシートを追加する
     *
     * @param sheetName シート名
     * @return 追加したシート
     */
    public BenriSheet createSheet(String sheetName) {
        Sheet sheet = wb.createSheet(sheetName);
        BenriSheet benriSheet = new BenriSheet(sheet);
        benriSheetMap.put(sheetName, benriSheet);
        return benriSheet;
    }
}
