package com.dwl.excel.style.configurer;

import com.dwl.excel.style.enums.BorderStyleValues;
import com.dwl.excel.style.enums.IndexedColorValues;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Objects;

public class BorderConfigurer {

  private BorderStyleValues topBorderStyle;
  private BorderStyleValues leftBorderStyle;
  private BorderStyleValues bottomBorderStyle;
  private BorderStyleValues rightBorderStyle;
  private boolean isTopBorder = false;
  private boolean isLeftBorder = false;
  private boolean isBottomBorder = false;
  private boolean isRightBorder = false;

  private IndexedColorValues topColor;
  private IndexedColorValues leftColor;
  private IndexedColorValues bottomColor;
  private IndexedColorValues rightColor;
  private boolean isTopColor = false;
  private boolean isLeftColor = false;
  private boolean isBottomColor = false;
  private boolean isRightColor = false;

  public BorderConfigurer styleAll(BorderStyleValues border){
    Objects.requireNonNull(border, "border can not be null");
    return styleTop(border)
        .styleLeft(border)
        .styleBottom(border)
        .styleRight(border);
  }

  public BorderConfigurer styleTop(BorderStyleValues borderStyle){
    Objects.requireNonNull(borderStyle, "borderStyle can not be null");
    this.isTopBorder = true;
    this.topBorderStyle = borderStyle;
    return this;
  }

  public BorderConfigurer styleLeft(BorderStyleValues borderStyle){
    Objects.requireNonNull(borderStyle, "border can not be null");
    this.isLeftBorder = true;
    this.leftBorderStyle = borderStyle;
    return this;
  }

  public BorderConfigurer styleBottom(BorderStyleValues borderStyle){
    Objects.requireNonNull(borderStyle, "border can not be null");
    this.isBottomBorder = true;
    this.bottomBorderStyle = borderStyle;
    return this;
  }

  public BorderConfigurer styleRight(BorderStyleValues borderStyle){
    Objects.requireNonNull(borderStyle, "borderStyle can not be null");
    this.isRightBorder = true;
    this.rightBorderStyle = borderStyle;
    return this;
  }

  public BorderConfigurer colorAll(IndexedColorValues color){
    return colorTop(color)
        .colorLeft(color)
        .colorBottom(color)
        .colorRight(color);
  }

  public BorderConfigurer colorTop(IndexedColorValues color){
    this.isTopColor = true;
    this.topColor = color;
    return this;
  }

  public BorderConfigurer colorLeft(IndexedColorValues color){
    this.isLeftColor = true;
    this.leftColor = color;
    return this;
  }

  public BorderConfigurer colorBottom(IndexedColorValues color){
    this.isBottomColor = true;
    this.bottomColor = color;
    return this;
  }

  public BorderConfigurer colorRight(IndexedColorValues color){
    this.isRightColor = true;
    this.rightColor = color;
    return this;
  }

  public BorderConfigurer apply(CellStyle cellStyle){
    Objects.requireNonNull(cellStyle, "cellStyle can not be null");
    if(isTopColor && topColor != null) cellStyle.setTopBorderColor(topColor.getPoiIndexedColors().getIndex());
    if(isLeftColor && leftColor != null) cellStyle.setLeftBorderColor(leftColor.getPoiIndexedColors().getIndex());
    if(isBottomColor && bottomColor != null) cellStyle.setBottomBorderColor(bottomColor.getPoiIndexedColors().getIndex());
    if(isRightColor && rightColor != null) cellStyle.setRightBorderColor(rightColor.getPoiIndexedColors().getIndex());
    if(isTopBorder && topBorderStyle != null) cellStyle.setBorderTop(topBorderStyle.getPoiBorderStyle());
    if(isLeftBorder && leftBorderStyle != null) cellStyle.setBorderLeft(leftBorderStyle.getPoiBorderStyle());
    if(isBottomBorder && bottomBorderStyle != null) cellStyle.setBorderBottom(bottomBorderStyle.getPoiBorderStyle());
    if(isRightBorder && rightBorderStyle != null) cellStyle.setBorderRight(rightBorderStyle.getPoiBorderStyle());
    return this;
  }
}