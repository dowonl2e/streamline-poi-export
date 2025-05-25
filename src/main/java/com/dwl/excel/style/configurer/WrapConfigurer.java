package com.dwl.excel.style.configurer;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Objects;

public class WrapConfigurer {

  private Boolean wrapText;

  public WrapConfigurer enable(){
    this.wrapText = true;
    return this;
  }

  public WrapConfigurer apply(CellStyle cellStyle){
    Objects.requireNonNull(cellStyle, "cellStyle can not be null");
    if(wrapText != null) cellStyle.setWrapText(wrapText);
    return this;
  }
}
