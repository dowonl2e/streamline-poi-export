package com.dwl.excel.style.configurer;

import com.dwl.excel.style.enums.HorizontalAlignmentValues;
import com.dwl.excel.style.enums.VerticalAlignmentValues;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Objects;

public final class AlignmentConfigurer {

  private HorizontalAlignmentValues horizontal;
  private VerticalAlignmentValues vertical;

  public AlignmentConfigurer horizontal(HorizontalAlignmentValues horizontal){
    this.horizontal = horizontal;
    return this;
  }

  public AlignmentConfigurer vertical(VerticalAlignmentValues vertical){
    this.vertical = vertical;
    return this;
  }

  public AlignmentConfigurer apply(CellStyle cellStyle){
    Objects.requireNonNull(cellStyle, "cellStyle can not be null");
    if(horizontal != null) cellStyle.setAlignment(horizontal.getPoiHorizontal());
    if(vertical != null) cellStyle.setVerticalAlignment(vertical.getPoiVertical());
    return this;
  }
}
