package com.monitorjbl.xlsx.impl;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormat;

public class StreamingDataFormat implements DataFormat {
  private StreamingStylesTable stylesSource;

  protected StreamingDataFormat(StreamingStylesTable stylesSource) {
      this.stylesSource = stylesSource;
  }

  /**
   * Get the format index that matches the given format string, creating a new
   * format entry if required. Aliases text to the proper format as required.
   *
   * @param format string matching a built in format
   * @return index of format.
   */
  public short getFormat(String format) {
    int idx = BuiltinFormats.getBuiltinFormat(format);
    if (idx == -1)
      idx = stylesSource.putNumberFormat(format);
    return (short) idx;
  }

  /**
   * get the format string that matches the given format index
   * 
   * @param index of a format
   * @return string represented at index of format or null if there is not a
   *         format at that index
   */
  public String getFormat(short index) {
    return getFormat(index & 0xffff);
  }

  /**
   * get the format string that matches the given format index
   * 
   * @param index of a format
   * @return string represented at index of format or null if there is not a
   *         format at that index
   */
  public String getFormat(int index) {
    String fmt = stylesSource.getNumberFormatAt(index);
    if (fmt == null)
      fmt = BuiltinFormats.getBuiltinFormat(index);
    return fmt;
  }

}
