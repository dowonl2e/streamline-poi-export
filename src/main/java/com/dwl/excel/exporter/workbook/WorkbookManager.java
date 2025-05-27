package com.dwl.excel.exporter.workbook;

import com.dwl.excel.style.CellStyleApplier;
import com.dwl.excel.style.font.FontStyleApplier;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;

public class WorkbookManager {

  private final Workbook workbook;

  public WorkbookManager(Workbook workbook){
    this.workbook = workbook;
  }

  public Sheet createSheet(String sheetName){
    return workbook.createSheet(sheetName);
  }

  public CellStyle createStyledCellStyle(CellStyleApplier cellStyleApplier){
    CellStyle cellStyle = this.workbook.createCellStyle();
    Font font = cellStyleApplier.getFontStyleApplier() != null
        ? createStyledFont(cellStyleApplier.getFontStyleApplier())
        : null;
    cellStyleApplier.apply(font, cellStyle);
    return cellStyle;
  }

  public Font createStyledFont(FontStyleApplier fontStyleApplier){
    Font font = this.workbook.createFont();
    fontStyleApplier.apply(font);
    return font;
  }

  public String getSheetName(int index){
    return this.workbook.getSheetName(index);
  }

  public int getCellStyleCount(){
    return this.workbook.getNumCellStyles();
  }

  public int getFontCount(){
    return this.workbook.getNumberOfFonts();
  }

  public void write(OutputStream stream) throws IOException {
    this.workbook.write(stream);
  }

  public void close() throws IOException {
    this.workbook.close();
  }
}
