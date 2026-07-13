package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.usermodel.Row;

public class StreamingEvaluationSheet implements EvaluationSheet {

  private final StreamingSheet sheet;

  public StreamingEvaluationSheet(StreamingSheet sheet) {
    this.sheet = sheet;
  }

  StreamingSheet getStreamingSheet() {
    return sheet;
  }

  /* Supported */

  @Override
  public EvaluationCell getCell(int rowIndex, int columnIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return null;
    }
    StreamingCell cell = (StreamingCell) row.getCell(columnIndex);
    if (cell == null) {
      return null;
    }
    return new StreamingEvaluationCell(cell, this);
  }

  @Override
  public void clearAllCachedResultValues() {
    return; // no caching is done, so nothing to clear
  }

  @Override
  public int getLastRowNum() {
    return sheet.getLastRowNum();
  }

  @Override
  public boolean isRowHidden(int rowIndex) {
    return false;
  }    
}
