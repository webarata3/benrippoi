# Benrippoi [![Build Status](https://travis-ci.org/webarata3/benrippoi.svg?branch=master)](https://travis-ci.org/webarata3/benrippoi) [![Coverage Status](https://coveralls.io/repos/github/webarata3/benrippoi/badge.svg?branch=master)](https://coveralls.io/github/webarata3/benrippoi?branch=master)


## サンプル

### ファイルの作成

```java
import link.webarata3.poi.BenriWorkbook;
import link.webarata3.poi.BenriWorkbookFactory;

import java.io.IOException;

public class CreateExcel {
    public static void main(String[] args) {
        BenriWorkbook bwb = BenriWorkbookFactory.createBlank();
        bwb.createSheet("ほげ").cell("A1").set("こんにちはExcel");
        try {
            bwb.write("hello.xlsx");
            bwb.close();
        } catch(IOException e) {
            e.printStackTrace();
        }
    }
}

```

### ファイルの読み込み

```java
import link.webarata3.poi.BenriSheet;
import link.webarata3.poi.BenriWorkbook;
import link.webarata3.poi.BenriWorkbookFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) {
        try (BenriWorkbook bwb = BenriWorkbookFactory.create("sample.xlsx")) {
            BenriSheet sheet = bwb.sheet("sheet1");
            System.out.println("A1=" + sheet.cell("A1").toStr());
            System.out.println("A2=" + sheet.cell("A2").toInt());
            System.out.println("A3=" + sheet.cell("A3").toDouble());
            System.out.println("A4=" + sheet.cell("A4").toBoolean());
            System.out.println("A5=" + sheet.cell("A5").toLocalDate());
            System.out.println("A6=" + sheet.cell("A6").toLocalTime());
            System.out.println("A7=" + sheet.cell("A7").toLocalDateTime());
        } catch(IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }
}
```

## ライセンス
MIT
