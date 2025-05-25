package com.dwl.excel.exporter.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.HashMap;
import java.util.Map;

public class StyleRegistry {

  private final CellStyle defaultHeaderCellStyle;
  private final CellStyle defaultBodyCellStyle;

  private final Map<String, CellStyle> cellStyleMap = new HashMap<>();
  private final Map<String, Font> fontMap = new HashMap<>();


  public StyleRegistry(
      CellStyle defaultHeaderCellStyle,
      CellStyle defaultBodyCellStyle
  ){
    this.defaultHeaderCellStyle = defaultHeaderCellStyle;
    this.defaultBodyCellStyle = defaultBodyCellStyle;
  }

  public CellStyle getDefaultHeaderCellStyle() {
    return defaultHeaderCellStyle;
  }

  public CellStyle getDefaultBodyCellStyle() {
    return defaultBodyCellStyle;
  }

  public Map<String, CellStyle> getCellStyleMap() {
    return this.cellStyleMap;
  }

  public Map<String, Font> getFontMap() {
    return this.fontMap;
  }

  public void addCellStyle(String key, CellStyle cellStyle){
    this.cellStyleMap.put(key, cellStyle);
  }

  public void addFont(String key, Font font){
    this.fontMap.put(key, font);
  }
}
