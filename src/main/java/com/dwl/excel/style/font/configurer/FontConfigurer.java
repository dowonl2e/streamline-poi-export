package com.dwl.excel.style.font.configurer;

import com.dwl.excel.style.enums.IndexedColorValues;
import com.dwl.excel.style.font.enums.FontCharsetValues;
import com.dwl.excel.style.font.enums.FontUnderlineValues;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Objects;

public class FontConfigurer {

  private IndexedColorValues color = IndexedColorValues.AUTOMATIC;
  private boolean isRGBA = false;
  private int red, green, blue, alpha;
  private Boolean bold;
  private FontUnderlineValues underline;
  private FontCharsetValues charset;
  private Short fontSize;
  private Boolean strikeOut;
  private Boolean italic;

  public FontConfigurer color(IndexedColorValues color){
    this.color = color;
    this.isRGBA = false;
    return this;
  }

  public FontConfigurer color(int red, int green, int blue){
    if(!((0 <= red && red <= 255) && (0 <= green && green <= 255) && (0 <= blue && blue <= 255))){
      throw new IllegalArgumentException("red, green, blue must be greater than or equal to 0 and less than or equal to 255.");
    }
    return color(red, green, blue,255);
  }

  public FontConfigurer color(int red, int green, int blue, int alpha){
    if(!((0 <= red && red <= 255) && (0 <= green && green <= 255)
        && (0 <= blue && blue <= 255) && (0 <= alpha && alpha <= 255))
    ){
      throw new IllegalArgumentException("red, green, blue, alpha must be greater than or equal to 0 and less than or equal to 255.");
    }
    this.isRGBA = true;
    this.red = red;
    this.green = green;
    this.blue = blue;
    this.alpha = alpha;
    return this;
  }

  public FontConfigurer bold(){
    this.bold = true;
    return this;
  }

  public FontConfigurer underline(FontUnderlineValues underline){
    this.underline = underline;
    return this;
  }

  public FontConfigurer size(short fontSize){
    this.fontSize = fontSize;
    return this;
  }

  public FontConfigurer sizeInPoints(short fontSize){
    if((fontSize*20) > Short.MAX_VALUE){
      throw new IllegalArgumentException("FontSize cannot have a value greater than "+(Short.MAX_VALUE/20));
    }
    this.fontSize = (short) (fontSize*20);
    return this;
  }

  public FontConfigurer charSet(FontCharsetValues charset){
    this.charset = charset;
    return this;
  }

  public FontConfigurer strikeOut(){
    this.strikeOut = true;
    return this;
  }

  public FontConfigurer italic(){
    this.italic = true;
    return this;
  }

  public FontConfigurer apply(Font font){
    Objects.requireNonNull(font, "font can not be null");
    if(isRGBA){
      XSSFColor color = new XSSFColor(new java.awt.Color(this.red, this.green, this.blue, this.alpha), null);
      ((XSSFFont)font).setColor(color);
    }
    else {
      font.setColor(color.getPoiIndexedColors().getIndex());
    }
    if(underline != null) font.setUnderline(underline.getPoiFontUnderline().getByteValue());
    if(bold != null) font.setBold(bold);
    if(fontSize != null) font.setFontHeight(fontSize);
    if(strikeOut != null) font.setStrikeout(strikeOut);
    if(italic != null) font.setItalic(italic);
    if(charset != null) font.setCharSet(charset.getPoiFontCharset().getNativeId());
    return this;
  }
}
