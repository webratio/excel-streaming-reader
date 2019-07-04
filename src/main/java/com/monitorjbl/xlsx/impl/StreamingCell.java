package com.monitorjbl.xlsx.impl;

import java.io.Serializable;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;

public class StreamingCell implements Cell, Serializable {
  private static final long serialVersionUID = 1L;

  private static final String FALSE_AS_STRING = "0";
  private static final String TRUE_AS_STRING  = "1";

  private int columnIndex;
  private int rowIndex;
  private final boolean use1904Dates;

  private Object contents;
  private Object rawContents;
  private String formula;
  private String numericFormat;
  private Short numericFormatIndex;
  private String type;
  private String cachedFormulaResultType;
  private Row row;
  private CellStyle cellStyle;
  private RichTextString richText;
  private Sheet sheet;

  public StreamingCell(StreamingSheet sheet, int columnIndex, int rowIndex, boolean use1904Dates) {
    this.columnIndex = columnIndex;
    this.rowIndex = rowIndex;
    this.use1904Dates = use1904Dates;
    this.sheet = sheet;
  }

  public Object getContents() {
    return contents;
  }

  public void setContents(Object contents) {
    this.contents = contents;
  }

  public Object getRawContents() {
    return rawContents;
  }

  public void setRawContents(Object rawContents) {
    this.rawContents = rawContents;
  }

  public String getNumericFormat() {
    return numericFormat;
  }

  public void setNumericFormat(String numericFormat) {
    this.numericFormat = numericFormat;
  }

  public Short getNumericFormatIndex() {
    return numericFormatIndex;
  }

  public void setNumericFormatIndex(Short numericFormatIndex) {
    this.numericFormatIndex = numericFormatIndex;
  }

  public void setFormula(String formula) {
    this.formula = formula;
  }

  public String getType() {
    return type;
  }

  public void setType(String type) {
    if("str".equals(type)) {
      // this is a formula cell, cache the value's type
      cachedFormulaResultType = this.type;
    }
    this.type = type;
  }

  public void setRow(Row row) {
    this.row = row;
  }

  @Override
  public void setCellStyle(CellStyle cellStyle) {
    this.cellStyle = cellStyle;
  }

  /* Supported */

  /**
   * Returns column index of this cell
   *
   * @return zero-based column index of a column in a sheet.
   */
  @Override
  public int getColumnIndex() {
    return columnIndex;
  }

  /**
   * Returns row index of a row in the sheet that contains this cell
   *
   * @return zero-based row index of a row in the sheet that contains this cell
   */
  @Override
  public int getRowIndex() {
    return rowIndex;
  }

  /**
   * Returns the Row this cell belongs to. Note that keeping references to cell
   * rows around after the iterator window has passed <b>will</b> preserve them.
   *
   * @return the Row that owns this cell
   */
  @Override
  public Row getRow() {
    return row;
  }

  /**
   * Return the cell type. Note that only the numeric, string, and blank types are
   * currently supported.
   *
   * @return the cell type
   * @throws UnsupportedOperationException Thrown if the type is not one supported by the streamer.
   *                                       It may be possible to still read the value as a supported type
   *                                       via {@code getStringCellValue()}, {@code getNumericCellValue},
   *                                       or {@code getDateCellValue()}
   * @see Cell#CELL_TYPE_BLANK
   * @see Cell#CELL_TYPE_NUMERIC
   * @see Cell#CELL_TYPE_STRING
   */
  @Override
  public int getCellType() {
    if(contents == null || type == null) {
      return CELL_TYPE_BLANK;
    } else if("n".equals(type)) {
      return CELL_TYPE_NUMERIC;
    } else if("s".equals(type) || "inlineStr".equals(type)) {
      return CELL_TYPE_STRING;
    } else if("str".equals(type)) {
      return CELL_TYPE_FORMULA;
    } else if("b".equals(type)) {
      return CELL_TYPE_BOOLEAN;
    } else if("e".equals(type)) {
      return CELL_TYPE_ERROR;
    } else {
      throw new UnsupportedOperationException("Unsupported cell type '" + type + "'");
    }
  }

  /**
   * Get the value of the cell as a string. For numeric cells we throw an exception.
   * For blank cells we return an empty string.
   *
   * @return the value of the cell as a string
   */
  @Override
  public String getStringCellValue() {
    return contents == null ? "" : (String) contents;
  }

  /**
   * Get the value of the cell as a number. For strings we throw an exception. For
   * blank cells we return a 0.
   *
   * @return the value of the cell as a number
   * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
   */
  @Override
  public double getNumericCellValue() {
    return rawContents == null ? 0.0 : Double.parseDouble((String) rawContents);
  }

  /**
   * Get the value of the cell as a date. For strings we throw an exception. For
   * blank cells we return a null.
   *
   * @return the value of the cell as a date
   * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
   * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
   */
  @Override
  public Date getDateCellValue() {
    if(getCellType() == CELL_TYPE_STRING){
      throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
    }
    return rawContents == null ? null : HSSFDateUtil.getJavaDate(getNumericCellValue(), use1904Dates);
  }

  /**
   * Get the value of the cell as a boolean. For strings we throw an exception. For
   * blank cells we return a false.
   *
   * @return the value of the cell as a date
   */
  @Override
  public boolean getBooleanCellValue() {
    int cellType = getCellType();
    switch(cellType) {
      case CELL_TYPE_BLANK:
        return false;
      case CELL_TYPE_BOOLEAN:
        return rawContents != null && TRUE_AS_STRING.equals(rawContents);
      case CELL_TYPE_FORMULA:
        throw new NotSupportedException();
      default:
        throw typeMismatch(CELL_TYPE_BOOLEAN, cellType, false);
    }
  }

  private static RuntimeException typeMismatch(int expectedTypeCode, int actualTypeCode, boolean isFormulaCell) {
    String msg = "Cannot get a "
            + getCellTypeName(expectedTypeCode) + " value from a "
            + getCellTypeName(actualTypeCode) + " " + (isFormulaCell ? "formula " : "") + "cell";
    return new IllegalStateException(msg);
  }

  /**
   * Used to help format error messages
   */
  private static String getCellTypeName(int cellTypeCode) {
    switch (cellTypeCode) {
      case CELL_TYPE_BLANK:   return "blank";
      case CELL_TYPE_STRING:  return "text";
      case CELL_TYPE_BOOLEAN: return "boolean";
      case CELL_TYPE_ERROR:   return "error";
      case CELL_TYPE_NUMERIC: return "numeric";
      case CELL_TYPE_FORMULA: return "formula";
    }
    return "#unknown cell type (" + cellTypeCode + ")#";
  }

  /**
   * @return the style of the cell
   */
  @Override
  public CellStyle getCellStyle() {
    return this.cellStyle;
  }

  /**
   * Return a formula for the cell, for example, <code>SUM(C4:E4)</code>
   *
   * @return a formula for the cell
   * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is not CELL_TYPE_FORMULA
   */
  @Override
  public String getCellFormula() {
    if (type == null || !"str".equals(type))
      throw new IllegalStateException("This cell does not have a formula");
    return formula;
  }

  /**
   * Only valid for formula cells
   * @return one of ({@link #CELL_TYPE_NUMERIC}, {@link #CELL_TYPE_STRING},
   *     {@link #CELL_TYPE_BOOLEAN}, {@link #CELL_TYPE_ERROR}) depending
   * on the cached value of the formula
   */
  @Override
  public int getCachedFormulaResultType() {
    if (type != null && "str".equals(type)) {
      if(contents == null || cachedFormulaResultType == null) {
        return CELL_TYPE_BLANK;
      } else if("n".equals(cachedFormulaResultType)) {
        return CELL_TYPE_NUMERIC;
      } else if("s".equals(cachedFormulaResultType) || "inlineStr".equals(cachedFormulaResultType)) {
        return CELL_TYPE_STRING;
      } else if("str".equals(cachedFormulaResultType)) {
        return CELL_TYPE_FORMULA;
      } else if("b".equals(cachedFormulaResultType)) {
        return CELL_TYPE_BOOLEAN;
      } else if("e".equals(cachedFormulaResultType)) {
        return CELL_TYPE_ERROR;
      } else {
        throw new UnsupportedOperationException("Unsupported cell type '" + cachedFormulaResultType + "'");
      }
    }
    else  {
      throw new IllegalStateException("Only formula cells have cached results");
    }
  }

  @Override
  public void setCellType(int cellType) {
    if (cellType == CELL_TYPE_BLANK) {
      setType(null);
    } else if (cellType == CELL_TYPE_NUMERIC) {
      setType("n");
    } else if (cellType == CELL_TYPE_STRING) {
      setType("s");
    } else if (cellType == CELL_TYPE_FORMULA) {
      setType("str");
    } else if (cellType == CELL_TYPE_BOOLEAN) {
      setType("b");
    } else if (cellType == CELL_TYPE_ERROR) {
      setType("e");
    } else {
      throw new UnsupportedOperationException("Unsupported cell type '" + cellType + "'");
    }
  }

  @Override
  public Sheet getSheet() {
    return this.sheet;
  }

  @Override
  public void setCellValue(double value) {
    if (Double.isInfinite(value)) {
      // Excel does not support positive/negative infinities,
      // rather, it gives a #DIV/0! error in these cases.
      setCellType(CELL_TYPE_ERROR);
      setContents(FormulaError.DIV0.getString());
    } else if (Double.isNaN(value)) {
      // Excel does not support Not-a-Number (NaN),
      // instead it immediately generates an #NUM! error.
      setCellType(CELL_TYPE_ERROR);
      setContents(FormulaError.NUM.getString());
    } else {
      setCellType(CELL_TYPE_NUMERIC);
      setContents(String.valueOf(value));
    }

  }

  @Override
  public void setCellValue(Date value) {
    setCellValue(DateUtil.getExcelDate(value, use1904Dates));
  }

  @Override
  public void setCellValue(Calendar value) {
    setCellValue(DateUtil.getExcelDate(value, use1904Dates));
  }

  @Override
  public void setCellValue(RichTextString value) {
    if (value == null || value.getString() == null) {
      setCellType(Cell.CELL_TYPE_BLANK);
      return;
    }

    if (value.length() > SpreadsheetVersion.EXCEL2007.getMaxTextLength()) {
      throw new IllegalArgumentException("The maximum length of cell contents (text) is 32,767 characters");
    }

    int cellType = getCellType();
    switch (cellType) {
    case Cell.CELL_TYPE_FORMULA:
      setContents(value.getString());
      setCellType(CELL_TYPE_STRING);
      break;
    default:
      setContents(value.getString());
      break;
    }

  }

  @Override
  public void setCellValue(String value) {
    setCellValue(value == null ? null : new XSSFRichTextString(value));
  }

  @Override
  public void setCellFormula(String formula) throws FormulaParseException {
    setFormula(formula);
  }

  void setRichText(RichTextString richText) {
    this.richText = richText;
  }

  @Override
  public RichTextString getRichStringCellValue() {
    if (this.richText == null) {
      this.richText = new XSSFRichTextString(getStringCellValue());
    }
    return this.richText;
  }

  @Override
  public void setCellValue(boolean value) {
    setCellType(CELL_TYPE_BOOLEAN);
    setRawContents(value);
  }

  /* Not supported */

  /**
   * Not supported
   */
  @Override
  public void setCellErrorValue(byte value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public byte getErrorCellValue() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setAsActiveCell() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellComment(Comment comment) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Comment getCellComment() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeCellComment() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Hyperlink getHyperlink() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setHyperlink(Hyperlink link) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeHyperlink() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public CellRangeAddress getArrayFormulaRange() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isPartOfArrayFormulaGroup() {
    throw new NotSupportedException();
  }
}
