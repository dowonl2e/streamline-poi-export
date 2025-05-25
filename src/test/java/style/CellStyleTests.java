package style;

import com.dwl.excel.style.CellStyleApplier;
import com.dwl.excel.style.enums.BorderStyleValues;
import com.dwl.excel.style.enums.FillPatternTypeValues;
import com.dwl.excel.style.enums.HorizontalAlignmentValues;
import com.dwl.excel.style.enums.IndexedColorValues;
import com.dwl.excel.style.enums.VerticalAlignmentValues;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class CellStyleTests {

  @Test
  @Order(1)
  @DisplayName("셀 테두리 스타일 테스트")
  public void cellStyleBorderTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .border(c ->
            c.styleTop(BorderStyleValues.THIN)
                .styleLeft(BorderStyleValues.MEDIUM)
                .styleBottom(BorderStyleValues.DOTTED)
                .styleRight(BorderStyleValues.THICK)
        )
        .build();

    cellStyler.apply(null, style);

    //then
    short expectedBorderTop = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedBorderLeft = BorderStyleValues.MEDIUM.getPoiBorderStyle().getCode();
    short expectedBorderBottom = BorderStyleValues.DOTTED.getPoiBorderStyle().getCode();
    short expectedBorderRight = BorderStyleValues.THICK.getPoiBorderStyle().getCode();

    Assertions.assertInstanceOf(XSSFCellStyle.class, style);
    Assertions.assertEquals(expectedBorderTop, style.getBorderTop().getCode());
    Assertions.assertEquals(expectedBorderLeft, style.getBorderLeft().getCode());
    Assertions.assertEquals(expectedBorderBottom, style.getBorderBottom().getCode());
    Assertions.assertEquals(expectedBorderRight, style.getBorderRight().getCode());
  }

  @Test
  @Order(2)
  @DisplayName("셀 테두리 색상 테스트")
  public void cellStyleBorderColorTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .border(c ->
            c.colorTop(IndexedColorValues.BLACK)
                .colorLeft(IndexedColorValues.RED)
                .colorBottom(IndexedColorValues.GREEN)
                .colorRight(IndexedColorValues.BLUE)
        )
        .build();

    cellStyler.apply(null, style);

    //then
    short expectedBorderTopColor = IndexedColorValues.BLACK.getPoiIndexedColors().getIndex();
    short expectedBorderLeftColor = IndexedColorValues.RED.getPoiIndexedColors().getIndex();
    short expectedBorderBottomColor = IndexedColorValues.GREEN.getPoiIndexedColors().getIndex();
    short expectedBorderRightColor = IndexedColorValues.BLUE.getPoiIndexedColors().getIndex();

    Assertions.assertInstanceOf(XSSFCellStyle.class, style);
    Assertions.assertEquals(expectedBorderTopColor, style.getTopBorderColor());
    Assertions.assertEquals(expectedBorderLeftColor, style.getLeftBorderColor());
    Assertions.assertEquals(expectedBorderBottomColor, style.getBottomBorderColor());
    Assertions.assertEquals(expectedBorderRightColor, style.getRightBorderColor());
  }

  @Test
  @Order(3)
  @DisplayName("셀 텍스트 배치 테스트")
  public void cellStyleAlignTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .alignment(c ->
            c.horizontal(HorizontalAlignmentValues.RIGHT)
                .vertical(VerticalAlignmentValues.BOTTOM)
        )
        .build();

    cellStyler.apply(null, style);

    //then
    short expectedHorizontalCode = HorizontalAlignmentValues.RIGHT.getPoiHorizontal().getCode();
    short expectedVerticalCode = VerticalAlignmentValues.BOTTOM.getPoiVertical().getCode();

    Assertions.assertInstanceOf(XSSFCellStyle.class, style);
    Assertions.assertEquals(expectedHorizontalCode, style.getAlignment().getCode());
    Assertions.assertEquals(expectedVerticalCode, style.getVerticalAlignment().getCode());
  }

  @Test
  @Order(4)
  @DisplayName("셀 배경색 테스트")
  public void cellStyleForegroundTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .foreground(c ->
            c.color(255,0,0, 240)
                .fill(FillPatternTypeValues.SOLID_FOREGROUND)
        )
        .build();

    cellStyler.apply(null, style);

    //then
    XSSFColor expectedRedColor = new XSSFColor(new java.awt.Color(255,0,0,240), null);
    short expectedFillPatternCode = FillPatternTypeValues.SOLID_FOREGROUND.getPoiFillPatternType().getCode();

    Assertions.assertInstanceOf(XSSFCellStyle.class, style);
    Assertions.assertEquals(expectedRedColor, style.getFillForegroundColorColor());
    Assertions.assertEquals(expectedFillPatternCode, style.getFillPattern().getCode());
  }

  @Test
  @Order(5)
  @DisplayName("셀 텍스트 감싸기 테스트")
  public void cellStyleWrapTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .wrap(c -> c.enable())
        .build();

    cellStyler.apply(null, style);

    //then
    Assertions.assertInstanceOf(XSSFCellStyle.class, style);
    Assertions.assertTrue(style.getWrapText());
  }

  @Test
  @Order(6)
  @DisplayName("셀 스타일 전체 테스트")
  public void cellStyleTotalTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    CellStyle style = workbook.createCellStyle();

    //when
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .alignment(c ->
            c.horizontal(HorizontalAlignmentValues.RIGHT)
                .vertical(VerticalAlignmentValues.BOTTOM)
        )
        .border(c ->
            c.styleTop(BorderStyleValues.THIN)
                .styleLeft(BorderStyleValues.MEDIUM)
                .styleBottom(BorderStyleValues.DOTTED)
                .styleRight(BorderStyleValues.THICK)
                .colorTop(IndexedColorValues.BLACK)
                .colorLeft(IndexedColorValues.RED)
                .colorBottom(IndexedColorValues.GREEN)
                .colorRight(IndexedColorValues.BLUE)
        )
        .foreground(c ->
            c.color(255,0,0,240).fill(FillPatternTypeValues.SOLID_FOREGROUND)
        )
        .wrap(c -> c.enable())
        .build();

    cellStyler.apply(null, style);

    //then
    Assertions.assertInstanceOf(XSSFCellStyle.class, style);

    short expectedBorderTop = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedBorderLeft = BorderStyleValues.MEDIUM.getPoiBorderStyle().getCode();
    short expectedBorderBottom = BorderStyleValues.DOTTED.getPoiBorderStyle().getCode();
    short expectedBorderRight = BorderStyleValues.THICK.getPoiBorderStyle().getCode();
    Assertions.assertEquals(expectedBorderTop, style.getBorderTop().getCode());
    Assertions.assertEquals(expectedBorderLeft, style.getBorderLeft().getCode());
    Assertions.assertEquals(expectedBorderBottom, style.getBorderBottom().getCode());
    Assertions.assertEquals(expectedBorderRight, style.getBorderRight().getCode());

    short expectedBorderTopColor = IndexedColorValues.BLACK.getPoiIndexedColors().getIndex();
    short expectedBorderLeftColor = IndexedColorValues.RED.getPoiIndexedColors().getIndex();
    short expectedBorderBottomColor = IndexedColorValues.GREEN.getPoiIndexedColors().getIndex();
    short expectedBorderRightColor = IndexedColorValues.BLUE.getPoiIndexedColors().getIndex();
    Assertions.assertEquals(expectedBorderTopColor, style.getTopBorderColor());
    Assertions.assertEquals(expectedBorderLeftColor, style.getLeftBorderColor());
    Assertions.assertEquals(expectedBorderBottomColor, style.getBottomBorderColor());
    Assertions.assertEquals(expectedBorderRightColor, style.getRightBorderColor());

    short expectedHorizontalCode = HorizontalAlignmentValues.RIGHT.getPoiHorizontal().getCode();
    short expectedVerticalCode = VerticalAlignmentValues.BOTTOM.getPoiVertical().getCode();
    Assertions.assertEquals(expectedHorizontalCode, style.getAlignment().getCode());
    Assertions.assertEquals(expectedVerticalCode, style.getVerticalAlignment().getCode());

    XSSFColor expectedRedColor = new XSSFColor(new java.awt.Color(255,0,0,240), null);
    short expectedFillPatternCode = FillPatternTypeValues.SOLID_FOREGROUND.getPoiFillPatternType().getCode();
    Assertions.assertEquals(expectedRedColor, style.getFillForegroundColorColor());
    Assertions.assertEquals(expectedFillPatternCode, style.getFillPattern().getCode());

    Assertions.assertTrue(style.getWrapText());
  }
}
