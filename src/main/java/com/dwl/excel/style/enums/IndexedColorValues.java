package com.dwl.excel.style.enums;

import org.apache.poi.ss.usermodel.IndexedColors;

public enum IndexedColorValues {
  BLACK1(IndexedColors.BLACK1),
  WHITE1(IndexedColors.WHITE1),
  RED1(IndexedColors.RED1),
  BRIGHT_GREEN1(IndexedColors.BRIGHT_GREEN1),
  BLUE1(IndexedColors.BLUE1),
  YELLOW1(IndexedColors.YELLOW1),
  PINK1(IndexedColors.PINK1),
  TURQUOISE1(IndexedColors.TURQUOISE1),
  BLACK(IndexedColors.BLACK),
  WHITE(IndexedColors.WHITE),
  RED(IndexedColors.RED),
  BRIGHT_GREEN(IndexedColors.BRIGHT_GREEN),
  BLUE(IndexedColors.BLUE),
  YELLOW(IndexedColors.YELLOW),
  PINK(IndexedColors.PINK),
  TURQUOISE(IndexedColors.TURQUOISE),
  DARK_RED(IndexedColors.DARK_RED),
  GREEN(IndexedColors.GREEN),
  DARK_BLUE(IndexedColors.DARK_BLUE),
  DARK_YELLOW(IndexedColors.DARK_YELLOW),
  VIOLET(IndexedColors.VIOLET),
  TEAL(IndexedColors.TEAL),
  GREY_25_PERCENT(IndexedColors.GREY_25_PERCENT),
  GREY_50_PERCENT(IndexedColors.GREY_50_PERCENT),
  CORNFLOWER_BLUE(IndexedColors.CORNFLOWER_BLUE),
  MAROON(IndexedColors.MAROON),
  LEMON_CHIFFON(IndexedColors.LEMON_CHIFFON),
  LIGHT_TURQUOISE1(IndexedColors.LIGHT_TURQUOISE1),
  ORCHID(IndexedColors.ORCHID),
  CORAL(IndexedColors.CORAL),
  ROYAL_BLUE(IndexedColors.ROYAL_BLUE),
  LIGHT_CORNFLOWER_BLUE(IndexedColors.LIGHT_CORNFLOWER_BLUE),
  SKY_BLUE(IndexedColors.SKY_BLUE),
  LIGHT_TURQUOISE(IndexedColors.LIGHT_TURQUOISE),
  LIGHT_GREEN(IndexedColors.LIGHT_GREEN),
  LIGHT_YELLOW(IndexedColors.LIGHT_YELLOW),
  PALE_BLUE(IndexedColors.PALE_BLUE),
  ROSE(IndexedColors.ROSE),
  LAVENDER(IndexedColors.LAVENDER),
  TAN(IndexedColors.TAN),
  LIGHT_BLUE(IndexedColors.LIGHT_BLUE),
  AQUA(IndexedColors.AQUA),
  LIME(IndexedColors.LIME),
  GOLD(IndexedColors.GOLD),
  LIGHT_ORANGE(IndexedColors.LIGHT_ORANGE),
  ORANGE(IndexedColors.ORANGE),
  BLUE_GREY(IndexedColors.BLUE_GREY),
  GREY_40_PERCENT(IndexedColors.GREY_40_PERCENT),
  DARK_TEAL(IndexedColors.DARK_TEAL),
  SEA_GREEN(IndexedColors.SEA_GREEN),
  DARK_GREEN(IndexedColors.DARK_GREEN),
  OLIVE_GREEN(IndexedColors.OLIVE_GREEN),
  BROWN(IndexedColors.BROWN),
  PLUM(IndexedColors.PLUM),
  INDIGO(IndexedColors.INDIGO),
  GREY_80_PERCENT(IndexedColors.GREY_80_PERCENT),
  AUTOMATIC(IndexedColors.AUTOMATIC);


  private final IndexedColors poiIndexedColors;

  IndexedColorValues(IndexedColors poiIndexedColors){
    this.poiIndexedColors = poiIndexedColors;
  }

  public IndexedColors getPoiIndexedColors() {
    return poiIndexedColors;
  }
}
