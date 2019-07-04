package com.monitorjbl.xlsx.impl.extensions;

import java.io.Serializable;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.util.Internal;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

public class StreamingCellBorder implements Serializable {

  private static final long serialVersionUID = 1L;

  private ThemesTable _theme;
  private CTBorder border;

  /**
   * Creates a Cell Border from the supplied XML definition
   * 
   * @param border the {@code CTBorder}
   * @param theme  the {@code ThemesTable}
   */
  public StreamingCellBorder(CTBorder border, ThemesTable theme) {
      this(border);
      this._theme = theme;
  }

  /**
   * Creates a Cell Border from the supplied XML definition
   * 
   * @param border the {@code CTBorder}
   */
  public StreamingCellBorder(CTBorder border) {
      this.border = border;
  }

  /**
   * Creates a new, empty Cell Border.
   * You need to attach this to the Styles Table
   */
  public StreamingCellBorder() {
      border = CTBorder.Factory.newInstance();
  }

  /**
   * Records the Themes Table that is associated with the current font, used when
   * looking up theme based colors and properties.
   * 
   * @param themes the {@code ThemesTable}
   */
  public void setThemesTable(ThemesTable themes) {
    this._theme = themes;
  }

  /**
   * The enumeration value indicating the side being used for a cell border.
   */
  public static enum BorderSide {
    TOP, RIGHT, BOTTOM, LEFT
  }

  /**
   * Returns the underlying XML bean.
   *
   * @return CTBorder
   */
  @Internal
  public CTBorder getCTBorder() {
    return border;
  }

  /**
   * Get the type of border to use for the selected border
   *
   * @param side - - where to apply the color definition
   * @return borderstyle - the type of border to use. default value is NONE if
   *         border style is not set.
   * @see BorderStyle
   */
  public BorderStyle getBorderStyle(BorderSide side) {
    CTBorderPr ctBorder = getBorder(side);
    STBorderStyle.Enum border = ctBorder == null ? STBorderStyle.NONE : ctBorder.getStyle();
    return BorderStyle.values()[border.intValue() - 1];
  }

  /**
   * Set the type of border to use for the selected border
   *
   * @param side  - - where to apply the color definition
   * @param style - border style
   * @see BorderStyle
   */
  public void setBorderStyle(BorderSide side, BorderStyle style) {
    getBorder(side, true).setStyle(STBorderStyle.Enum.forInt(style.ordinal() + 1));
  }

  /**
   * Get the color to use for the selected border
   *
   * @param side - where to apply the color definition
   * @return color - color to use as XSSFColor. null if color is not set
   */
  public XSSFColor getBorderColor(BorderSide side) {
    CTBorderPr borderPr = getBorder(side);

    if (borderPr != null && borderPr.isSetColor()) {
      XSSFColor clr = new XSSFColor(borderPr.getColor());
      if (_theme != null) {
        _theme.inheritFromThemeAsRequired(clr);
      }
      return clr;
    } else {
      // No border set
      return null;
    }
  }

  /**
   * Set the color to use for the selected border
   *
   * @param side  - where to apply the color definition
   * @param color - the color to use
   */
  public void setBorderColor(BorderSide side, XSSFColor color) {
    CTBorderPr borderPr = getBorder(side, true);
    if (color == null)
      borderPr.unsetColor();
    else
      borderPr.setColor(color.getCTColor());
  }

  private CTBorderPr getBorder(BorderSide side) {
    return getBorder(side, false);
  }


  private CTBorderPr getBorder(BorderSide side, boolean ensure) {
    CTBorderPr borderPr;
    switch (side) {
    case TOP:
      borderPr = border.getTop();
      if (ensure && borderPr == null)
        borderPr = border.addNewTop();
      break;
    case RIGHT:
      borderPr = border.getRight();
      if (ensure && borderPr == null)
        borderPr = border.addNewRight();
      break;
    case BOTTOM:
      borderPr = border.getBottom();
      if (ensure && borderPr == null)
        borderPr = border.addNewBottom();
      break;
    case LEFT:
      borderPr = border.getLeft();
      if (ensure && borderPr == null)
        borderPr = border.addNewLeft();
      break;
    default:
      throw new IllegalArgumentException("No suitable side specified for the border");
    }
    return borderPr;
  }


  public int hashCode() {
    return border.toString().hashCode();
  }

  public boolean equals(Object o) {
    if (!(o instanceof StreamingCellBorder))
      return false;

    StreamingCellBorder cf = (StreamingCellBorder) o;
    return border.toString().equals(cf.getCTBorder().toString());
  }
}
