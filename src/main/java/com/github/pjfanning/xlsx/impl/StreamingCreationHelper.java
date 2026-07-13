package com.github.pjfanning.xlsx.impl;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public class StreamingCreationHelper implements CreationHelper {

  private final StreamingWorkbook workbook;

  public StreamingCreationHelper(StreamingWorkbook workbook) {
    this.workbook = workbook;
  }

  @Override
  public RichTextString createRichTextString(String text) {
    throw new UnsupportedOperationException("Unimplemented method 'createRichTextString'");
  }

  @Override
  public DataFormat createDataFormat() {
    throw new UnsupportedOperationException("Unimplemented method 'createDataFormat'");
  }

  @Override
  public Hyperlink createHyperlink(HyperlinkType type) {
    throw new UnsupportedOperationException("Unimplemented method 'createHyperlink'");
  }

  @Override
  public FormulaEvaluator createFormulaEvaluator() {
    return new StreamingFormulaEvaluator(workbook);
  }

  @Override
  public ExtendedColor createExtendedColor() {
    throw new UnsupportedOperationException("Unimplemented method 'createExtendedColor'");
  }

  @Override
  public ClientAnchor createClientAnchor() {
    throw new UnsupportedOperationException("Unimplemented method 'createClientAnchor'");
  }

  @Override
  public AreaReference createAreaReference(String reference) {
    throw new UnsupportedOperationException("Unimplemented method 'createAreaReference'");
  }

  @Override
  public AreaReference createAreaReference(CellReference topLeft, CellReference bottomRight) {
    throw new UnsupportedOperationException("Unimplemented method 'createAreaReference'");
  }

}
