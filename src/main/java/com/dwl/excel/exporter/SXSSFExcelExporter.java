package com.dwl.excel.exporter;

import com.dwl.excel.style.CellStyleApplier;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;
import java.util.Optional;

public class SXSSFExcelExporter extends ExcelExporter {
  private final Integer flushCount;

  private SXSSFExcelExporter(
      Integer flushCount, int rowAccessWindowSize,
      boolean compressTmpFiles, boolean useSharedStringsTable,
      CellStyleApplier defaultHeaderCellStyleApplier, CellStyleApplier defaultBodyCellStyleApplier
  ) {
    super(
        new SXSSFWorkbook(new XSSFWorkbook(), rowAccessWindowSize, compressTmpFiles, useSharedStringsTable),
        SpreadsheetVersion.EXCEL2007,
        Optional.ofNullable(defaultHeaderCellStyleApplier),
        Optional.ofNullable(defaultBodyCellStyleApplier)
    );
    this.flushCount = flushCount;
  }

  public static Builder builder(){
    return new Builder();
  }

  public static class Builder {

    private Integer flushCount;
    private int rowAccessWindowSize;
    private boolean compressTmpFiles;
    private boolean useSharedStringsTable;
    private CellStyleApplier defaultHeaderCellStyleApplier;
    private CellStyleApplier defaultBodyCellStyleApplier;
    public Builder(){
      flushCount = 100;
      rowAccessWindowSize = -1;
      compressTmpFiles = false;
      useSharedStringsTable = false;
    }

    public Builder flushCount(Integer flushCount){
      this.flushCount = flushCount;
      return this;
    }

    public Builder rowAccessWindowSize(int rowAccessWindowSize){
      this.rowAccessWindowSize = rowAccessWindowSize;
      return this;
    }

    public Builder compressTmpFiles(boolean compressTmpFiles){
      this.compressTmpFiles = compressTmpFiles;
      return this;
    }

    public Builder useSharedStringsTable(boolean useSharedStringsTable){
      this.useSharedStringsTable = useSharedStringsTable;
      return this;
    }

    public Builder defaultHeaderCellStyleApplier(CellStyleApplier defaultHeaderCellStyleApplier){
      this.defaultHeaderCellStyleApplier = defaultHeaderCellStyleApplier;
      return this;
    }

    public Builder defaultBodyCellStyleApplier(CellStyleApplier defaultBodyCellStyleApplier){
      this.defaultBodyCellStyleApplier = defaultBodyCellStyleApplier;
      return this;
    }

    public SXSSFExcelExporter build(){
      return new SXSSFExcelExporter(
          flushCount, rowAccessWindowSize,
          compressTmpFiles, useSharedStringsTable,
          defaultHeaderCellStyleApplier, defaultBodyCellStyleApplier
      );
    }
  }

  @Override
  protected void flush() throws Exception {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int flushNumber = currentRow%flushCount;
    if(flushNumber == 0){
      ((SXSSFSheet)sheet).flushRows(flushCount);
    }
  }
}
