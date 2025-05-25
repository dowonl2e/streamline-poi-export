package com.dwl.excel.style.enums;

import org.apache.poi.ss.usermodel.VerticalAlignment;

public enum VerticalAlignmentValues {
  TOP(VerticalAlignment.TOP),
  CENTER(VerticalAlignment.CENTER),
  BOTTOM(VerticalAlignment.BOTTOM),
  JUSTIFY(VerticalAlignment.JUSTIFY),
  DISTRIBUTED(VerticalAlignment.DISTRIBUTED);

  private final VerticalAlignment poiVertical;

  VerticalAlignmentValues(VerticalAlignment poiVertical){
    this.poiVertical = poiVertical;
  }

  public VerticalAlignment getPoiVertical() {
    return poiVertical;
  }
}
