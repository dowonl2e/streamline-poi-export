package com.dwl.excel.exporter.writer;

import com.dwl.excel.MaximumExceededException;
import com.dwl.excel.NotAllowedClassException;
import com.dwl.excel.annotation.ExportTargetField;
import com.dwl.excel.exporter.Exportable;
import com.dwl.excel.exporter.style.StyleRegistry;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

public final class CellWriter {

  private final String DEFAULT_STYLE_KEY = "DEFAULT_STYLE";
  private final Sheet sheet;

  private final CellStyle defaultStyle;
  private final Map<String, CellStyle> styleMap;
  private final Map<String, Font> fontMap;

  private final int currentRow;
  private final int maxColumns;
  private final int maxTextSize;
  private int currentCol = 0;

  public CellWriter(
      Sheet sheet,
      boolean isHeader,
      StyleRegistry styleRegistry,
      int currentRow,
      int maxColumns,
      int maxTextSize
  ){
    this.sheet = sheet;
    this.defaultStyle = isHeader
        ? styleRegistry.getDefaultHeaderCellStyle()
        : styleRegistry.getDefaultBodyCellStyle();
    this.styleMap = styleRegistry.getCellStyleMap();
    this.fontMap = styleRegistry.getFontMap();
    this.currentRow = currentRow;
    this.currentCol = findStartColumn();
    this.maxColumns = maxColumns;
    this.maxTextSize = maxTextSize;
  }

  private Row getOrCreateRow(int rowIndex){
    return this.sheet.getRow(rowIndex) == null ? this.sheet.createRow(rowIndex) : this.sheet.getRow(rowIndex);
  }

  private int findStartColumn(){
    Row row = getOrCreateRow(this.currentRow);
    for (int i = 0, last = row.getLastCellNum(); i <= last; i++) {
      if (row.getCell(i) == null) return i;
    }
    return 0;
  }

  public CellWriter nextCol(){
    this.currentCol++;
    return this;
  }

  public CellWriter nextCol(int add){
    if(add <= 0) {
      throw new IllegalArgumentException("Column cannot move to the current or previous column (add > 0)");
    }
    this.currentCol += add;
    return this;
  }

  public int getCurrentCol(){
    return this.currentCol;
  }

  private CellStyle getOrDefaultStyle(String styleKey){
    return this.styleMap.getOrDefault(styleKey, this.defaultStyle);
  }

  public CellWriter writeTargetHeaders(Class<? extends Exportable> exportableClass){
    return writeTargetHeaders(DEFAULT_STYLE_KEY, exportableClass);
  }

  public CellWriter writeTargetHeaders(String styleKey, Class<? extends Exportable> exportableClass){
    List<String> headers = extractHeaders(exportableClass);
    CellStyle style = getOrDefaultStyle(styleKey);
    headers.forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeTargetCells(Exportable objectValue) {
    return writeTargetCells(DEFAULT_STYLE_KEY, objectValue);
  }

  public CellWriter writeTargetCells(String styleKey, Exportable objectValue) {
    List<?> values = extractValues(objectValue);
    CellStyle style = getOrDefaultStyle(styleKey);
    values.forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  private List<String> extractHeaders(Class<? extends Exportable> exportableClass){
    Field[] fields = exportableClass.getDeclaredFields();

    List<ExportTargetField> targetAnnotations = Arrays.stream(fields)
        .filter(field -> field.isAnnotationPresent(ExportTargetField.class))
        .map(field -> field.getAnnotation(ExportTargetField.class))
        .sorted(Comparator.comparing(ExportTargetField::order))
        .collect(Collectors.toList());

    return targetAnnotations.stream().map(ExportTargetField::header).collect(Collectors.toList());
  }

  private List<?> extractValues(Exportable objectValue) {
    if(objectValue == null) return new ArrayList<>();

    Field[] fields = objectValue.getClass().getDeclaredFields();

    return Arrays.stream(fields)
        .filter(field -> field.isAnnotationPresent(ExportTargetField.class))
        .sorted(Comparator.comparing(f -> f.getAnnotation(ExportTargetField.class).order()))
        .map(field -> {
          try {
            return invokeGetter(objectValue, field);
          } catch (Exception e) {
            throw new RuntimeException(e);
          }
        })
        .collect(Collectors.toList());
  }

  private Object invokeGetter(Object object, Field field) throws Exception {
    String fieldName = field.getName();
    Class<?> clazz = object.getClass();

    String capitalized = Character.toUpperCase(fieldName.charAt(0)) + fieldName.substring(1);
    boolean isBoolean = field.getType() == boolean.class || field.getType() == Boolean.class;
    String getterName = isBoolean ? "is" + capitalized : "get" + capitalized;

    try {
      Method getter = clazz.getMethod(getterName);
      return getter.invoke(object);
    } catch (NoSuchMethodException e) {
      String message = "Could not find getter for field '" + fieldName + "'";
      if(isBoolean) message += " Boolean(boolean) type method names must be started with 'is'";
      throw new NoSuchMethodException(message);
    } catch (Exception e) {
      throw new Exception("Getter invocation failed '" + fieldName + "'", e);
    }
  }

  public CellWriter writeCell(String value){
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, String value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCell(String value, boolean isFormula){
    return writeCell(DEFAULT_STYLE_KEY, value, isFormula);
  }

  public CellWriter writeCell(String styleKey, String value, boolean isFormula){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value, isFormula);
    return this;
  }

  public CellWriter writeCells(String[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, String[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(Long value){
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, Long value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(Long[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, Long[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(Double value){
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, Double value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(Double[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, Double[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(Integer value){
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, Integer value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(Integer[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, Integer[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(Boolean value){
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, Boolean value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(Boolean[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, Boolean[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(Date value) {
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, Date value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(Date[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, Date[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(LocalDate value) {
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, LocalDate value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(LocalDate[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, LocalDate[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  public CellWriter writeCell(LocalDateTime value) {
    return writeCell(DEFAULT_STYLE_KEY, value);
  }

  public CellWriter writeCell(String styleKey, LocalDateTime value){
    CellStyle style = getOrDefaultStyle(styleKey);
    writeAndFillMergedRegion(style, value);
    return this;
  }

  public CellWriter writeCells(LocalDateTime[] values){
    return writeCells(DEFAULT_STYLE_KEY, values);
  }

  public CellWriter writeCells(String styleKey, LocalDateTime[] values){
    CellStyle style = getOrDefaultStyle(styleKey);
    List.of(values).forEach(value -> writeAndFillMergedRegion(style, value));
    return this;
  }

  private Optional<CellRangeAddress> findCellRangeAddress(int row, int col) {
    return this.sheet.getMergedRegions()
        .stream()
        .filter(cellAddresses -> cellAddresses.isInRange(row, col))
        .findFirst();
  }

  private void writeAndFillMergedRegion(CellStyle style, Object value){
    writeAndFillMergedRegion(style, value, false);
  }

  private void writeAndFillMergedRegion(CellStyle style, Object value, boolean isFormula){
    Optional<CellRangeAddress> optional = findCellRangeAddress(this.currentRow, this.currentCol);
    optional.ifPresentOrElse(range -> {
      int startRow = range.getFirstRow();
      int endRow = range.getLastRow();
      int startCol = range.getFirstColumn();
      int endCol = range.getLastColumn();

      for(int i = startRow ; i <= endRow ; i++){
        for(int j = startCol ; j <= endCol ; j++){
          Object cellValue = (i == startRow && j == startCol) ? value : null;
          writeCell(style, i, j, cellValue, isFormula);
        }
      }
    }, () -> writeCell(style, this.currentRow, this.currentCol, value, isFormula));
  }

  private void writeCell(CellStyle cellStyle, int rowIndex, int colIndex, Object value, boolean isFormula) {
    if (this.currentCol > maxColumns) {
      throw new MaximumExceededException("Number of columns is exceeded (limit: " + maxColumns + ")");
    }
    if (value != null && value.toString().length() > maxTextSize) {
      throw new MaximumExceededException("Length of value is exceeded (limit: " + maxTextSize + ")");
    }

    Row row = getOrCreateRow(rowIndex);
    Cell cell;
    if (value == null) {
      cell = row.createCell(colIndex, CellType.BLANK);
    } else if (value instanceof Long) {
      cell = row.createCell(colIndex, CellType.NUMERIC);
      cell.setCellValue((Long) value);
    } else if (value instanceof String) {
      CellType strCellType = isFormula ? CellType.FORMULA : CellType.STRING;
      cell = row.createCell(colIndex, strCellType);
      cell.setCellValue((String) value);
    } else if (value instanceof Double) {
      cell = row.createCell(colIndex, CellType.NUMERIC);
      cell.setCellValue((Double) value);
    } else if (value instanceof Integer) {
      cell = row.createCell(colIndex, CellType.NUMERIC);
      cell.setCellValue((Integer) value);
    } else if (value instanceof Boolean) {
      cell = row.createCell(colIndex, CellType.BOOLEAN);
      cell.setCellValue((Boolean) value);
    } else if (value instanceof Date) {
      cell = row.createCell(colIndex, CellType.STRING);
      cell.setCellValue(value.toString());
    } else if (value instanceof LocalDate) {
      cell = row.createCell(colIndex, CellType.STRING);
      cell.setCellValue(value.toString());
    } else if (value instanceof LocalDateTime) {
      cell = row.createCell(colIndex, CellType.STRING);
      cell.setCellValue(value.toString());
    } else {
      throw new NotAllowedClassException("This type is not allowed for extraction '" + (value.getClass().getTypeName()) + "'");
    }

    if (cellStyle != null){
      cell.setCellStyle(cellStyle);
    }

    this.currentCol = colIndex+1;
  }

  private Font getFont(String fontKey){
    return this.fontMap.get(fontKey);
  }

  public CellWriter writeRichTextCell(String fontKey, String value, int startIndex, int endIndex){
    writeRichTextCell(DEFAULT_STYLE_KEY, fontKey, value, startIndex, endIndex);
    return this;
  }

  public CellWriter writeRichTextCell(String styleKey, String fontKey, String value, int startIndex, int endIndex){
    CellStyle style = getOrDefaultStyle(styleKey);
    Font font = getFont(fontKey);
    writeRichTextAndFillMergedRegion(style, font, value, startIndex, endIndex);
    return this;
  }

  private void writeRichTextAndFillMergedRegion(
      CellStyle style, Font font, String value,
      int startIndex, int endIndex
  ){
    Optional<CellRangeAddress> optional = findCellRangeAddress(this.currentRow, this.currentCol);
    optional.ifPresentOrElse(range -> {
      int startRow = range.getFirstRow();
      int endRow = range.getLastRow();
      int startCol = range.getFirstColumn();
      int endCol = range.getLastColumn();

      for(int i = startRow ; i <= endRow ; i++){
        for(int j = startCol ; j <= endCol ; j++){
          String cellValue = (i == startRow && j == startCol) ? value : null;
          writeRichTextCell(style, font, i, j, cellValue, startIndex, endIndex);
        }
      }
    }, () -> writeRichTextCell(style, font, this.currentRow, this.currentCol, value, startIndex, endIndex));
  }

  private void writeRichTextCell(
      CellStyle style, Font font,
      int rowIndex, int colIndex, String value,
      int startIndex, int endIndex) {
    if (this.currentCol > maxColumns) {
      throw new MaximumExceededException("Number of columns is exceeded (limit: " + maxColumns + ")");
    }
    if (value != null && value.length() > maxTextSize) {
      throw new MaximumExceededException("Length of value is exceeded (limit: " + maxTextSize + ")");
    }

    Row row = getOrCreateRow(rowIndex);

    CellType cellType = value == null ? CellType.BLANK : CellType.STRING;
    Cell cell = row.createCell(colIndex, cellType);

    if(value != null) {
      RichTextString richText = cell instanceof HSSFCell
          ? new HSSFRichTextString(value)
          : new XSSFRichTextString(value);

      if(font != null) {
        richText.applyFont(startIndex, endIndex, font);
      }
      if(style != null) {
        cell.setCellStyle(style);
      }
      cell.setCellValue(richText);
    }

    this.currentCol = colIndex+1;
  }
}
