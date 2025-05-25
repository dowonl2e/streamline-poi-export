package com.dwl.excel.style.configurer;

import com.dwl.excel.style.enums.FillPatternTypeValues;
import com.dwl.excel.style.enums.IndexedColorValues;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.Objects;

public class ForegroundConfigurer {

  private IndexedColorValues color = IndexedColorValues.AUTOMATIC;

  private boolean isRGBA = false;

  private int red, green, blue, alpha;

  private FillPatternTypeValues fillPatternType;

  public ForegroundConfigurer fill(FillPatternTypeValues fillPatternType){
    Objects.requireNonNull(fillPatternType, "fillPatternType can not be null");
    this.fillPatternType = fillPatternType;
    return this;
  }

  public ForegroundConfigurer color(IndexedColorValues color){
    fill(FillPatternTypeValues.SOLID_FOREGROUND);
    this.color = color;
    this.isRGBA = false;
    return this;
  }

  public ForegroundConfigurer color(int red, int green, int blue){
    if(!((0 <= red && red <= 255) && (0 <= green && green <= 255) && (0 <= blue && blue <= 255))){
      throw new IllegalArgumentException("red, green, blue must be greater than or equal to 0 and less than or equal to 255.");
    }
    return color(red, green, blue, 255);
  }

  public ForegroundConfigurer color(int red, int green, int blue, int alpha){
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

  public ForegroundConfigurer apply(CellStyle cellStyle){
    Objects.requireNonNull(cellStyle, "cellStyle can not be null");
    if(isRGBA){
       Color color = new XSSFColor(new java.awt.Color(this.red, this.green, this.blue, this.alpha), null);
       cellStyle.setFillForegroundColor(color);
    }
    else {
      cellStyle.setFillForegroundColor(color.getPoiIndexedColors().getIndex());
    }
    if(fillPatternType != null) cellStyle.setFillPattern(fillPatternType.getPoiFillPatternType());

    return this;
  }
}
