package com.monitorjbl.xlsx.impl;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;

public class StreamingEvaluationCell implements EvaluationCell {

  private final EvaluationSheet evalSheet;
  private final StreamingCell cell;

  public StreamingEvaluationCell(StreamingCell cell, StreamingEvaluationSheet evaluationSheet) {
    this.cell = cell;
    this.evalSheet = evaluationSheet;
  }

  public StreamingEvaluationCell(StreamingCell cell) {
    this(cell, new StreamingEvaluationSheet((StreamingSheet) cell.getSheet()));
  }

  /* Supported */

  @Override
  public Object getIdentityKey() {
    // save memory by just using the cell itself as the identity key
    // Note - this assumes XSSFCell has not overridden hashCode and equals
    return cell;
  }

  public StreamingCell getStreamingCell() {
    return cell;
  }

  @Override
  public boolean getBooleanCellValue() {
    return cell.getBooleanCellValue();
  }

  @Override
  public int getCellType() {
    return cell.getCellType();
  }

  @Override
  public int getColumnIndex() {
    return cell.getColumnIndex();
  }

  @Override
  public int getErrorCellValue() {
    return cell.getErrorCellValue();
  }

  @Override
  public double getNumericCellValue() {
    return cell.getNumericCellValue();
  }

  @Override
  public int getRowIndex() {
    return cell.getRowIndex();
  }

  @Override
  public EvaluationSheet getSheet() {
    return evalSheet;
  }

  @Override
  public String getStringCellValue() {
    return cell.getRichStringCellValue().getString();
  }

  @Override
  public int getCachedFormulaResultType() {
    return cell.getCachedFormulaResultType();
  }

}
