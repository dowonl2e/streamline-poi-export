package com.dwl.excel.exporter;

import com.dwl.excel.MaximumExceededException;
import com.dwl.excel.exporter.sheet.SheetConfigurer;
import com.dwl.excel.exporter.style.StyleRegistry;
import com.dwl.excel.exporter.workbook.WorkbookBuilder;
import com.dwl.excel.exporter.writer.CellWriter;
import com.dwl.excel.style.CellStyleApplier;
import com.dwl.excel.style.font.FontStyleApplier;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;

public abstract class ExcelExporter {

  private static final String DEFAULT_SHEET_NAME = "Sheet";
  private final SpreadsheetVersion version;

  protected Workbook workbook;
  protected Sheet sheet;

  private final WorkbookBuilder workbookBuilder;
  private final StyleRegistry styleRegistry;
  private SheetConfigurer sheetConfigurer;

  protected int currentRow = 0;

  protected ExcelExporter(
      Workbook workbook,
      SpreadsheetVersion spreadsheetVersion,
      Optional<CellStyleApplier> defaultHeaderCellStyleApplier,
      Optional<CellStyleApplier> defaultBodyCellStyleApplier
  ){
    Objects.requireNonNull(workbook, "workbook must not be null");

    this.workbook = workbook;
    this.version = spreadsheetVersion;
    this.workbookBuilder = new WorkbookBuilder(this.workbook);
    this.styleRegistry = new StyleRegistry(
        defaultHeaderCellStyleApplier.map(workbookBuilder::createStyledCellStyle).orElse(null),
        defaultBodyCellStyleApplier.map(workbookBuilder::createStyledCellStyle).orElse(null)
    );
  }

  public SheetConfigurer createSheet(){
    return createSheet(DEFAULT_SHEET_NAME+getSheetCount()+1);
  }

  public SheetConfigurer createSheet(String sheetName){
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    this.sheet = this.workbookBuilder.createSheet(sheetName);
    this.sheetConfigurer = new SheetConfigurer(this.sheet);
    this.currentRow = 0;
    return this.sheetConfigurer;
  }

  public ExcelExporter createCellStyle(String styleKey, CellStyleApplier cellStyleApplier){
    Objects.requireNonNull(styleKey, "styleKey must not be null");
    Objects.requireNonNull(cellStyleApplier, "cellStyleApplier must not be null");
    CellStyle cellStyle = this.workbookBuilder.createStyledCellStyle(cellStyleApplier);
    this.styleRegistry.addCellStyle(styleKey, cellStyle);
    return this;
  }

  public ExcelExporter createFont(String fontKey, FontStyleApplier fontStyler){
    Objects.requireNonNull(fontKey, "fontKey must not be null");
    Objects.requireNonNull(fontStyler, "fontStyler must not be null");
    Font font = this.workbookBuilder.createStyledFont(fontStyler);
    this.styleRegistry.addFont(fontKey, font);
    return this;
  }

  public int getSheetCount(){
    return this.workbook.getNumberOfSheets();
  }

  public String getSheetName(int index){
    return this.workbookBuilder.getSheetName(index);
  }

  public int getCellStyleCount(){
    return this.workbookBuilder.getCellStyleCount();
  }

  public int getFontCount(){
    return this.workbookBuilder.getFontCount();
  }

  public ExcelExporter nextRow(){
    this.currentRow++;
    return this;
  }

  public ExcelExporter nextRow(int add){
    if(add <= 0) {
      throw new IllegalArgumentException("Row can not move to the current or previous row. (add <= 0)");
    }
    this.currentRow += add;
    return this;
  }

  public int getCurrentRow(){
    return this.currentRow;
  }

  public ExcelExporter mergeCell(int startRow, int endRow, int startColumn, int endColumn){
    Objects.requireNonNull(this.sheetConfigurer, "sheetConfigurator must not be null");
    this.sheetConfigurer.mergeCell(startRow, endRow, startColumn, endColumn);
    return this;
  }

  public CellWriter createHeader() throws Exception {
    return createRow(true);
  }

  public CellWriter createRow() throws Exception {
    return createRow(false);
  }

  private CellWriter createRow(boolean isHeader) throws Exception {
    Objects.requireNonNull(this.sheet, "sheet must not be null");
    if(this.currentRow > this.version.getMaxRows()){
      throw new MaximumExceededException("number of rows in the sheet is exceeded (limit: " + version.getMaxRows() +")");
    }

    CellWriter writer = new CellWriter(
        this.sheet, isHeader, this.styleRegistry,
        this.currentRow, this.version.getMaxColumns(), this.version.getMaxTextLength()
    );

    this.currentRow++;
    flush();

    return writer;
  }

  protected abstract void flush() throws Exception;

  public void export(OutputStream stream) throws IOException{
    Objects.requireNonNull(stream, "stream must not be null");
    this.workbookBuilder.write(stream);
    this.workbookBuilder.close();
    stream.flush();
    stream.close();
  }
}
