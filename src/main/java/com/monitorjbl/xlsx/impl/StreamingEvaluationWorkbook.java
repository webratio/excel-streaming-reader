package com.monitorjbl.xlsx.impl;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationName;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.SheetIdentifier;
import org.apache.poi.ss.formula.ptg.NamePtg;
import org.apache.poi.ss.formula.ptg.NameXPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public class StreamingEvaluationWorkbook implements EvaluationWorkbook, FormulaParsingWorkbook {

  private final StreamingWorkbook workbook;

  public StreamingEvaluationWorkbook(StreamingWorkbook workbook) {
    this.workbook = workbook;
  }

  /* Supported */

  @Override
  public Ptg[] getFormulaTokens(EvaluationCell evalCell) {
    StreamingCell cell = ((StreamingEvaluationCell) evalCell).getStreamingCell();
    StreamingEvaluationWorkbook frBook = new StreamingEvaluationWorkbook(workbook);
    return FormulaParser.parse(cell.getCellFormula(), frBook, FormulaType.CELL,
        workbook.getSheetIndex(cell.getSheet()));
  }

  @Override
  public EvaluationSheet getSheet(int sheetIndex) {
    return new StreamingEvaluationSheet((StreamingSheet) workbook.getSheetAt(sheetIndex));
  }

  @Override
  public int getSheetIndex(EvaluationSheet evalSheet) {
    StreamingSheet sheet = ((StreamingEvaluationSheet) evalSheet).getStreamingSheet();
    return workbook.getSheetIndex(sheet);
  }

  @Override
  public String getSheetName(int sheetIndex) {
    return workbook.getSheetName(sheetIndex);
  }

  @Override
  public int getSheetIndex(String sheetName) {
    return workbook.getSheetIndex(sheetName);
  }

  @Override
  public UDFFinder getUDFFinder() {
    return null; // allowed value
  }

  @Override
  public SpreadsheetVersion getSpreadsheetVersion() {
    return SpreadsheetVersion.EXCEL2007;
  }

  /* Not supported */

  @Override
  public ExternalSheet getExternalSheet(int externSheetIndex) {
    throw new UnsupportedOperationException();
  }

  @Override
  public ExternalSheet getExternalSheet(String firstSheetName, String lastSheetName, int externalWorkbookNumber) {
    throw new UnsupportedOperationException();
  }

  @Override
  public int convertFromExternSheetIndex(int externSheetIndex) {
    throw new UnsupportedOperationException();
  }

  @Override
  public ExternalName getExternalName(int externSheetIndex, int externNameIndex) {
    throw new UnsupportedOperationException();
  }

  @Override
  public ExternalName getExternalName(String nameName, String sheetName, int externalWorkbookNumber) {
    throw new UnsupportedOperationException();
  }

  @Override
  public EvaluationName getName(NamePtg namePtg) {
    throw new UnsupportedOperationException();
  }

  @Override
  public EvaluationName getName(String name, int sheetIndex) {
    throw new UnsupportedOperationException();
  }

  @Override
  public String resolveNameXText(NameXPtg ptg) {
    throw new UnsupportedOperationException();
  }

  @Override
  public Ptg getNameXPtg(String name, SheetIdentifier sheet) {
    throw new UnsupportedOperationException();
  }

  @Override
  public Ptg get3DReferencePtg(CellReference cell, SheetIdentifier sheet) {
    throw new UnsupportedOperationException();
  }

  @Override
  public Ptg get3DReferencePtg(AreaReference area, SheetIdentifier sheet) {
    throw new UnsupportedOperationException();
  }

  @Override
  public int getExternalSheetIndex(String sheetName) {
    throw new UnsupportedOperationException();
  }

  @Override
  public int getExternalSheetIndex(String workbookName, String sheetName) {
    throw new UnsupportedOperationException();
  }

}
