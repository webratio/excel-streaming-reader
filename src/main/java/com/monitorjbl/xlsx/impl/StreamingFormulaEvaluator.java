package com.monitorjbl.xlsx.impl;

import java.util.Map;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class StreamingFormulaEvaluator implements FormulaEvaluator {

  private final StreamingWorkbook workbook;
  private final WorkbookEvaluator workbookEvaluator;

  public StreamingFormulaEvaluator(StreamingWorkbook workbook) {
    this.workbook = workbook;
    this.workbookEvaluator = new WorkbookEvaluator(new StreamingEvaluationWorkbook(workbook), null, null);
  }

  /* Supported */

  @Override
  public void clearAllCachedResultValues() {
    workbookEvaluator.clearAllCachedResultValues();
  }

  @Override
  public Cell evaluateInCell(Cell cell) {
    if (cell == null) {
      return cell;
    }
    if (cell.getCellType() == Cell.CELL_TYPE_FORMULA && !isFormulaEmpty(cell)) {
      CellValue cv = evaluateFormulaCellValue(cell);
      setCellType(cell, cv); // cell will no longer be a formula cell
      setCellValue(cell, cv);
    }
    return cell;
  }

  /**
   * Not supported
   */
  @Override
  public void notifySetFormula(Cell cell) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void notifyDeleteCell(Cell cell) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void notifyUpdateCell(Cell cell) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void evaluateAll() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellValue evaluate(Cell cell) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int evaluateFormulaCell(Cell cell) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setupReferencedWorkbooks(Map<String, FormulaEvaluator> workbooks) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setIgnoreMissingWorkbooks(boolean ignore) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setDebugEvaluationOutputForNextEval(boolean value) {
    throw new UnsupportedOperationException();
  }

  /* Private */

  private static void setCellType(Cell cell, CellValue cv) {
    int cellType = cv.getCellType();
    switch (cellType) {
    case Cell.CELL_TYPE_BOOLEAN:
    case Cell.CELL_TYPE_ERROR:
    case Cell.CELL_TYPE_NUMERIC:
    case Cell.CELL_TYPE_STRING:
      cell.setCellType(cellType);
      return;
    case Cell.CELL_TYPE_BLANK:
      // never happens - blanks eventually get translated to zero
    case Cell.CELL_TYPE_FORMULA:
      // this will never happen, we have already evaluated the formula
    }
    throw new IllegalStateException("Unexpected cell value type (" + cellType + ")");
  }

  private static void setCellValue(Cell cell, CellValue cv) {
    int cellType = cv.getCellType();
    switch (cellType) {
    case Cell.CELL_TYPE_BOOLEAN:
      cell.setCellValue(cv.getBooleanValue());
      break;
    case Cell.CELL_TYPE_ERROR:
      cell.setCellErrorValue(cv.getErrorValue());
      break;
    case Cell.CELL_TYPE_NUMERIC:
      cell.setCellValue(cv.getNumberValue());
      break;
    case Cell.CELL_TYPE_STRING:
      cell.setCellValue(new XSSFRichTextString(cv.getStringValue()));
      break;
    case Cell.CELL_TYPE_BLANK:
      // never happens - blanks eventually get translated to zero
    case Cell.CELL_TYPE_FORMULA:
      // this will never happen, we have already evaluated the formula
    default:
      throw new IllegalStateException("Unexpected cell value type (" + cellType + ")");
    }
  }

  /**
   * Returns a CellValue wrapper around the supplied ValueEval instance.
   */
  private CellValue evaluateFormulaCellValue(Cell cell) {
    EvaluationCell evalCell = toEvaluationCell(cell);
    ValueEval eval = workbookEvaluator.evaluate(evalCell);
    if (eval instanceof NumberEval) {
      NumberEval ne = (NumberEval) eval;
      return new CellValue(ne.getNumberValue());
    }
    if (eval instanceof BoolEval) {
      BoolEval be = (BoolEval) eval;
      return CellValue.valueOf(be.getBooleanValue());
    }
    if (eval instanceof StringEval) {
      StringEval ne = (StringEval) eval;
      return new CellValue(ne.getStringValue());
    }
    if (eval instanceof ErrorEval) {
      return CellValue.getError(((ErrorEval) eval).getErrorCode());
    }
    throw new RuntimeException("Unexpected eval class (" + eval.getClass().getName() + ")");
  }

  private EvaluationCell toEvaluationCell(Cell cell) {
    if (!(cell instanceof StreamingCell)) {
      throw new IllegalArgumentException(
          "Unexpected type of cell: " + cell.getClass() + "." + " Only StreamingCell can be evaluated.");
    }

    return new StreamingEvaluationCell((StreamingCell) cell);
  }

  private boolean isFormulaEmpty(Cell cell) {
    if (cell.getCellFormula() == null) {
      return true;
    }
    return cell.getCellFormula().isEmpty();
  }
}
