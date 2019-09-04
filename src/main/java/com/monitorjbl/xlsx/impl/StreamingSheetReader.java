package com.monitorjbl.xlsx.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import com.monitorjbl.xlsx.exceptions.CloseException;

public class StreamingSheetReader implements Iterable<Row>, Serializable {

  private static final long serialVersionUID = 1L;

  private static final Logger log = LoggerFactory.getLogger(StreamingSheetReader.class);

  private final SharedStringsTable sst;
  private final StreamingStylesTable stylesTable;
  private final transient XMLEventReader parser;
  private final StreamingDataFormatter dataFormatter = new StreamingDataFormatter();
  private final Set<Integer> hiddenColumns = new HashSet<>();

  private int lastRowNum;
  private int rowCacheSize;
  private List<Row> rowCache;
  private Iterator<Row> rowCacheIterator;

  private String lastContents;
  private StreamingRow currentRow;
  private StreamingCell currentCell;
  private StreamingSheet sheet;
  private boolean use1904Dates;
  private final Map<Integer, File> rowsByRownum = new HashMap<Integer, File>();
  private final Set<String> rowsFiles = new HashSet<String>();
  private File currentRowFile;
  private RowWindow currentRowFileWindow;

  public StreamingSheetReader(SharedStringsTable sst, StreamingStylesTable stylesTable, XMLEventReader parser,
      final boolean use1904Dates, int rowCacheSize) {
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.parser = parser;
    this.use1904Dates = use1904Dates;
    this.rowCacheSize = rowCacheSize;
    this.currentRowFileWindow = null;
    this.rowCache = new ArrayList<Row>(rowCacheSize);
  }

  /**
   * Read through a number of rows equal to the rowCacheSize field or until there is no more data to read
   *
   * @return true if data was read
   */
  private boolean getRow() {
    try {
      rowCache.clear();
      while(rowCache.size() < rowCacheSize && parser.hasNext()) {
        handleEvent(parser.nextEvent());
      }
      rowCacheIterator = rowCache.iterator();
      boolean somethingRead = rowCacheIterator.hasNext();
      if (somethingRead) {
        currentRowFile = getRowsFileName();
        int minRowNum = Integer.MAX_VALUE;
        int maxRowNum = Integer.MIN_VALUE;
        int[] rowMap = new int[rowCacheSize];
        for (Row row : rowCache) {
          int currentRowNum = row.getRowNum();
          rowsByRownum.put(currentRowNum, currentRowFile);
          minRowNum = Math.min(minRowNum, currentRowNum);
          maxRowNum = Math.max(maxRowNum, currentRowNum);
        }

        // iterate again the row cache to build the mapping between the row num and the
        // actual position in the array
        int rowMapIndex = 0;
        for (Row row : rowCache) {
          int currentRowNum = row.getRowNum();
          rowMap[currentRowNum - minRowNum] = rowMapIndex++;
        }

        currentRowFileWindow = new RowWindow(minRowNum, maxRowNum, rowMap);
        writeRows(currentRowFile, currentRowFileWindow);
      }
      return somethingRead;
    } catch(XMLStreamException | SAXException e) {
      log.debug("End of stream");
    }
    return false;
  }

  /**
   * Handles a SAX event.
   *
   * @param event
   * @throws SAXException
   */
  private void handleEvent(XMLEvent event) throws SAXException {
    if(event.getEventType() == XMLStreamConstants.CHARACTERS) {
      Characters c = event.asCharacters();
      lastContents += c.getData();
    } else if(event.getEventType() == XMLStreamConstants.START_ELEMENT
        && isSpreadsheetTag(event.asStartElement().getName())) {
      StartElement startElement = event.asStartElement();
      String tagLocalName = startElement.getName().getLocalPart();

      if("row".equals(tagLocalName)) {
        Attribute rowNumAttr = startElement.getAttributeByName(new QName("r"));
        Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
        int rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
        boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
        currentRow = new StreamingRow(rowIndex, isHidden);
      } else if("col".equals(tagLocalName)) {
        Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
        boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
        if(isHidden) {
          Attribute minAttr = startElement.getAttributeByName(new QName("min"));
          Attribute maxAttr = startElement.getAttributeByName(new QName("max"));
          int min = Integer.parseInt(minAttr.getValue()) - 1;
          int max = Integer.parseInt(maxAttr.getValue()) - 1;
          for(int columnIndex = min; columnIndex <= max; columnIndex++)
            hiddenColumns.add(columnIndex);
        }
      } else if("c".equals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(new QName("r"));

        String[] coord = ref.getValue().split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
        currentCell = new StreamingCell(sheet, CellReference.convertColStringToIndex(coord[0]),
            Integer.parseInt(coord[1]) - 1, use1904Dates);
        setFormatString(startElement, currentCell);

        Attribute type = startElement.getAttributeByName(new QName("t"));
        if(type != null) {
          currentCell.setType(type.getValue());
        } else {
          currentCell.setType("n");
        }

        Attribute style = startElement.getAttributeByName(new QName("s"));
        if(style != null) {
          String indexStr = style.getValue();
          try {
            int index = Integer.parseInt(indexStr);
            currentCell.setCellStyle(stylesTable.getStyleAt(index));
          } catch(NumberFormatException nfe) {
            log.warn("Ignoring invalid style index {}", indexStr);
          }
        }
      } else if("dimension".equals(tagLocalName)) {
        Attribute refAttr = startElement.getAttributeByName(new QName("ref"));
        String ref = refAttr != null ? refAttr.getValue() : null;
        if(ref != null) {
          // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
          for(int i = ref.length() - 1; i >= 0; i--) {
            if(!Character.isDigit(ref.charAt(i))) {
              try {
                lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
              } catch(NumberFormatException ignore) { }
              break;
            }
          }
        }
      } else if("f".equals(tagLocalName)) {
        currentCell.setType("str");
      }

      // Clear contents cache
      lastContents = "";
    } else if(event.getEventType() == XMLStreamConstants.END_ELEMENT
        && isSpreadsheetTag(event.asEndElement().getName())) {
      EndElement endElement = event.asEndElement();
      String tagLocalName = endElement.getName().getLocalPart();

      if("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
        currentCell.setRawContents(unformattedContents());
        currentCell.setContents(formattedContents());
      } else if("row".equals(tagLocalName) && currentRow != null) {
        rowCache.add(currentRow);
      } else if("c".equals(tagLocalName)) {
        currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
      } else if("f".equals(tagLocalName)) {
        currentCell.setFormula(lastContents);
      }

    }
  }

  /**
   * Returns true if a tag is part of the main namespace for SpreadsheetML:
   * <ul>
   * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
   * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
   * </ul>
   * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
   *
   * @param name
   * @return
   */
  private boolean isSpreadsheetTag(QName name) {
    return (name.getNamespaceURI() != null
        && name.getNamespaceURI().endsWith("/main"));
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  boolean isColumnHidden(int columnIndex) {
    if(rowCacheIterator == null) {
      getRow();
    }
    return hiddenColumns.contains(columnIndex);
  }

  /**
   * Gets the last row on the sheet
   *
   * @return
   */
  int getLastRowNum() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return lastRowNum;
  }

  /**
   * Read the numeric format string out of the styles table for this cell. Stores
   * the result in the Cell.
   *
   * @param startElement
   * @param cell
   */
  void setFormatString(StartElement startElement, StreamingCell cell) {
    Attribute cellStyle = startElement.getAttributeByName(new QName("s"));
    String cellStyleString = (cellStyle != null) ? cellStyle.getValue() : null;
    StreamingCellStyle style = null;

    if(cellStyleString != null) {
      style = stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
    } else if(stylesTable.getNumCellStyles() > 0) {
      style = stylesTable.getStyleAt(0);
    }

    if(style != null) {
      cell.setNumericFormatIndex(style.getDataFormat());
      String formatString = style.getDataFormatString();

      if(formatString != null) {
        cell.setNumericFormat(formatString);
      } else {
        cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
      }
    } else {
      cell.setNumericFormatIndex(null);
      cell.setNumericFormat(null);
    }
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the type of cell and the discovered numeric format.
   *
   * @return
   */
  String formattedContents() {
    switch(currentCell.getType()) {
      case "s":           //string stored in shared table
        int idx = Integer.parseInt(lastContents);
        return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      case "inlineStr":   //inline string (not in sst)
        return new XSSFRichTextString(lastContents).toString();
      case "str":         //forumla type
        return '"' + lastContents + '"';
      case "e":           //error type
        return "ERROR:  " + lastContents;
      case "n":           //numeric type
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          return dataFormatter.formatRawCellContents(
              Double.parseDouble(lastContents),
              currentCell.getNumericFormatIndex(),
              currentCell.getNumericFormat());
        } else {
          return lastContents;
        }
      default:
        return lastContents;
    }
  }

  RichTextString richTextString() {
    switch (currentCell.getType()) {
    case "s": // string stored in shared table
      int idx = Integer.parseInt(lastContents);
      return new XSSFRichTextString(sst.getEntryAt(idx));
    case "inlineStr": // inline string (not in sst)
      return new XSSFRichTextString(lastContents);
    case "str": // forumla type
      return new XSSFRichTextString('"' + lastContents + '"');
    case "e": // error type
      return new XSSFRichTextString("ERROR:  " + lastContents);
    case "n": // numeric type
      if (currentCell.getNumericFormat() != null && lastContents.length() > 0) {
        return new XSSFRichTextString(dataFormatter.formatRawCellContents(Double.parseDouble(lastContents),
            currentCell.getNumericFormatIndex(), currentCell.getNumericFormat()));
      } else {
        return new XSSFRichTextString(lastContents);
      }
    default:
      return new XSSFRichTextString(lastContents);
    }
  }

  /**
   * Returns the contents of the cell, with no formatting applied
   *
   * @return
   */
  String unformattedContents() {
    switch(currentCell.getType()) {
      case "s":           //string stored in shared table
        int idx = Integer.parseInt(lastContents);
        return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      case "inlineStr":   //inline string (not in sst)
        return new XSSFRichTextString(lastContents).toString();
      default:
        return lastContents;
    }
  }

  /**
   * Returns a new streaming iterator to loop through rows. This iterator is not
   * guaranteed to have all rows in memory, and any particular iteration may
   * trigger a load from disk to read in new data.
   *
   * @return the streaming iterator
   */
  @Override
  public Iterator<Row> iterator() {
    return new StreamingRowIterator();
  }

  public void close() {
    try {
      parser.close();
      // clean up the files
      for (String rowsFile : rowsFiles) {
        File f = new File(rowsFile);
        log.debug("Deleting rows file [" + f.getAbsolutePath() + "]");
        f.delete();
      }
      rowsFiles.clear();
      rowsByRownum.clear();
    } catch(XMLStreamException e) {
      throw new CloseException(e);
    }
  }

  Row getRow(int rownum) {
    if (currentRowFileWindow == null || !currentRowFileWindow.contains(rownum)) {
      // makes the reader to read some more rows
      getRow();
    }
    // if the wanted row did't get read, it means there is no such a row (empty row)
    if (currentRowFileWindow.contains(rownum) && !rowsByRownum.containsKey(rownum)) {
      return null;
    }
    // read the row bunch from the file
    File rowsFile = rowsByRownum.get(rownum);
    if (!rowsFile.getAbsolutePath().equals(currentRowFile.getAbsolutePath())) {
      // we do not have in memory the wanted cells, so we have to load them
      currentRowFileWindow = readRows(rowsFile);
      currentRowFile = rowsFile;
      rowsFiles.add(rowsFile.getAbsolutePath());
    }
    int rowCacheIndex = currentRowFileWindow.getRowCacheIndex(rownum);
    return rowCache.get(rowCacheIndex);
  }

  private File getRowsFileName() {
    File rowsFile = null;
    try {
      String rowsFileName = Files.createTempFile(String.valueOf(System.currentTimeMillis()), ".ser").toString();
      rowsFile = new File(rowsFileName);
      rowsFile.mkdirs();
    } catch (IOException e) {
      e.printStackTrace();
    }
    return rowsFile;
  }

  private void writeRows(File rowsFile, RowWindow rowWindow) {
    log.debug("Writing rows to file [" + rowsFile.getAbsolutePath() + "]");
    try (FileOutputStream fout = new FileOutputStream(rowsFile);
        ObjectOutputStream oos = new ObjectOutputStream(fout);) {
      oos.writeObject(rowWindow);
      oos.writeObject(rowCache);
      rowsFiles.add(rowsFile.getAbsolutePath());
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private RowWindow readRows(File rowsFile) {
    log.debug("Reading rows from file [" + rowsFile.getAbsolutePath() + "]");
    rowCache.clear();
    RowWindow rowWindow = null;
    try (FileInputStream fin = new FileInputStream(rowsFile); ObjectInputStream ois = new ObjectInputStream(fin);) {
      rowWindow = (RowWindow) ois.readObject();
      rowCache = (List<Row>) ois.readObject();
    } catch (Exception e) {
      e.printStackTrace();
    }
    return rowWindow;
  }

  void setSheet(StreamingSheet sheet) {
    this.sheet = sheet;
  }

  StreamingSheet getSheet() {
    return this.sheet;
  }

  class StreamingRowIterator implements Iterator<Row> {
    public StreamingRowIterator() {
      if(rowCacheIterator == null) {
        hasNext();
      }
    }

    @Override
    public boolean hasNext() {
      return (rowCacheIterator != null && rowCacheIterator.hasNext()) || getRow();
    }

    @Override
    public Row next() {
      return rowCacheIterator.next();
    }

    @Override
    public void remove() {
      throw new RuntimeException("NotSupported");
    }
  }

  private class RowWindow implements Serializable {
    private static final long serialVersionUID = 1L;

    private final int minRowNum;
    private final int maxRowNum;
    private final int[] rowMap;

    RowWindow(int minRowNum, int maxRowNum, int[] rowMap) {
      this.minRowNum = minRowNum;
      this.maxRowNum = maxRowNum;
      this.rowMap = rowMap;
    }

    boolean contains(int rowNum) {
      return minRowNum <= rowNum && maxRowNum >= rowNum;
    }

    int getRowCacheIndex(int rowNum) {
      return rowMap[rowNum - minRowNum];
    }
  }

  private void writeObject(ObjectOutputStream os) throws IOException {
  }

  private void readObject(ObjectInputStream is) throws IOException, ClassNotFoundException {
  }
}
