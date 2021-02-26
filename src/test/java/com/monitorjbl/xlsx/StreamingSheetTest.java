package com.monitorjbl.xlsx;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.BeforeClass;
import org.junit.Test;

public class StreamingSheetTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testLastRowNum() throws Exception {
    try (InputStream is = new FileInputStream(new File("src/test/resources/large.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(24, sheet.getLastRowNum());
    }

    try (InputStream is = new FileInputStream(new File("src/test/resources/empty_sheet.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getLastRowNum());
    }

    try (InputStream is = new FileInputStream(new File("src/test/resources/with_rownum.xlsx"));
        Workbook workbook = StreamingReader.builder().rowCacheSize(10).bufferSize(4096).open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(14216, sheet.getLastRowNum());
    }

    try (InputStream is = new FileInputStream(new File("src/test/resources/without_rownum.xlsx"));
        Workbook workbook = StreamingReader.builder().rowCacheSize(10).bufferSize(4096).open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(9, sheet.getLastRowNum());
    }

    try (InputStream is = new FileInputStream(new File("src/test/resources/row_spans_over_multiple_columns.xlsx"));
        Workbook workbook = StreamingReader.builder().rowCacheSize(1000).bufferSize(4096).open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(28, sheet.getLastRowNum());
    }

    try (InputStream is = new FileInputStream(new File("src/test/resources/from_libreoffice.xlsx"));
        Workbook workbook = StreamingReader.builder().rowCacheSize(1000).bufferSize(4096).open(is);) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(1048575, sheet.getLastRowNum());
    }
  }

}
