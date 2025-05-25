package com.dwl.excel.exporter;

import com.dwl.excel.style.CellStyleApplier;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Optional;

public class XSSFExcelExporter extends ExcelExporter {

  private XSSFExcelExporter(
      CellStyleApplier defaultHeaderCellStyleApplier,
      CellStyleApplier defaultBodyCellStyleApplier
  ) {
    super(
        new XSSFWorkbook(),
        SpreadsheetVersion.EXCEL2007,
        Optional.ofNullable(defaultHeaderCellStyleApplier),
        Optional.ofNullable(defaultBodyCellStyleApplier)
    );
  }

  public static XSSFExcelExporter.Builder builder(){
    return new XSSFExcelExporter.Builder();
  }

  public static class Builder {

    private CellStyleApplier defaultHeaderCellStyleApplier;
    private CellStyleApplier defaultBodyCellStyleApplier;
    public Builder(){}

    public XSSFExcelExporter.Builder defaultHeaderCellStyleApplier(CellStyleApplier defaultHeaderCellStyleApplier){
      this.defaultHeaderCellStyleApplier = defaultHeaderCellStyleApplier;
      return this;
    }

    public XSSFExcelExporter.Builder defaultBodyCellStyleApplier(CellStyleApplier defaultBodyCellStyleApplier){
      this.defaultBodyCellStyleApplier = defaultBodyCellStyleApplier;
      return this;
    }

    public XSSFExcelExporter build(){
      return new XSSFExcelExporter(defaultHeaderCellStyleApplier, defaultBodyCellStyleApplier);
    }
  }

  @Override
  protected void flush() throws Exception {
    /* Do Nothing */
  }

}
