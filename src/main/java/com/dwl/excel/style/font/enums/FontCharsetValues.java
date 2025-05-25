package com.dwl.excel.style.font.enums;

import org.apache.poi.common.usermodel.fonts.FontCharset;

public enum FontCharsetValues {

  ANSI(FontCharset.ANSI),
  DEFAULT(FontCharset.DEFAULT),
  SYMBOL(FontCharset.SYMBOL),
  MAC(FontCharset.MAC),
  SHIFTJIS(FontCharset.SHIFTJIS),
  HANGUL(FontCharset.HANGUL),
  JOHAB(FontCharset.JOHAB),
  GB2312(FontCharset.GB2312),
  CHINESEBIG5(FontCharset.CHINESEBIG5),
  GREEK(FontCharset.GREEK),
  TURKISH(FontCharset.TURKISH),
  VIETNAMESE(FontCharset.VIETNAMESE),
  HEBREW(FontCharset.HEBREW),
  ARABIC(FontCharset.ARABIC),
  BALTIC(FontCharset.BALTIC),
  RUSSIAN(FontCharset.RUSSIAN),
  THAI(FontCharset.THAI),
  EASTEUROPE(FontCharset.EASTEUROPE),
  OEM(FontCharset.OEM);

  private final FontCharset poiFontCharset;

  FontCharsetValues(FontCharset poiFontCharset) {
    this.poiFontCharset = poiFontCharset;
  }

  public FontCharset getPoiFontCharset() {
    return poiFontCharset;
  }
}
