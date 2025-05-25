package com.dwl.excel.style.font.enums;

import org.apache.poi.ss.usermodel.FontUnderline;

public enum FontUnderlineValues {

  SINGLE(FontUnderline.SINGLE),
  DOUBLE(FontUnderline.DOUBLE),
  SINGLE_ACCOUNTING(FontUnderline.SINGLE_ACCOUNTING),
  DOUBLE_ACCOUNTING(FontUnderline.DOUBLE_ACCOUNTING),
  NONE(FontUnderline.NONE);

  private final FontUnderline poiFontUnderline;

  FontUnderlineValues(FontUnderline poiFontUnderline) {
    this.poiFontUnderline = poiFontUnderline;
  }

  public FontUnderline getPoiFontUnderline() {
    return poiFontUnderline;
  }
}
