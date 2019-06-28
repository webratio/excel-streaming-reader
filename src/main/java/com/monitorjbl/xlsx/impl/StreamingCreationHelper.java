package com.monitorjbl.xlsx.impl;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;

public class StreamingCreationHelper implements CreationHelper {

  private final StreamingWorkbook workbook;

  public StreamingCreationHelper(StreamingWorkbook workbook) {
    this.workbook = workbook;
  }

  @Override
  public RichTextString createRichTextString(String text) {
    throw new UnsupportedOperationException();
  }

  @Override
  public DataFormat createDataFormat() {
    throw new UnsupportedOperationException();
  }

  @Override
  public Hyperlink createHyperlink(int type) {
    throw new UnsupportedOperationException();
  }

  @Override
  public FormulaEvaluator createFormulaEvaluator() {
    return new StreamingFormulaEvaluator(workbook);
  }

  @Override
  public ExtendedColor createExtendedColor() {
    throw new UnsupportedOperationException();
  }

  @Override
  public ClientAnchor createClientAnchor() {
    throw new UnsupportedOperationException();
  }

}
