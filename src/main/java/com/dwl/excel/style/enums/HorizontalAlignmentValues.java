package com.dwl.excel.style.enums;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public enum HorizontalAlignmentValues {
  GENERAL(HorizontalAlignment.GENERAL),
  LEFT(HorizontalAlignment.LEFT),
  CENTER(HorizontalAlignment.CENTER),
  RIGHT(HorizontalAlignment.RIGHT),
  FILL(HorizontalAlignment.FILL),
  JUSTIFY(HorizontalAlignment.JUSTIFY),
  CENTER_SELECTION(HorizontalAlignment.CENTER_SELECTION),
  DISTRIBUTED(HorizontalAlignment.DISTRIBUTED);

  private final HorizontalAlignment poiHorizontal;
  HorizontalAlignmentValues(HorizontalAlignment poiHorizontal){
    this.poiHorizontal = poiHorizontal;
  }

  public HorizontalAlignment getPoiHorizontal(){
    return this.poiHorizontal;
  }
}
