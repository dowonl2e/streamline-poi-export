package sheet;

import com.dwl.excel.exporter.sheet.SheetConfigurer;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;

public class SheetConfiguratorTests {

  @Test
  @Order(1)
  @DisplayName("시트 세부 설정 테스트")
  public void sheetOptionsTest(){
    //given
    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet();
    SheetConfigurer configurator = new SheetConfigurer(sheet);

    boolean autoBreaks = sheet.getAutobreaks();
    int defaultColumnWidth = sheet.getDefaultColumnWidth();
    short defaultRowHeight = sheet.getDefaultRowHeight();
    boolean isRightToLeft = sheet.isRightToLeft();
    boolean isRowSumsRight = sheet.getRowSumsRight();
    boolean isRowSumsBellow = sheet.getRowSumsBelow();
    boolean isSelected = sheet.isSelected();

    //when
    configurator.autoBreaks(!autoBreaks);
    configurator.activeCell(0, 0);
    configurator.columnWidth(0, defaultColumnWidth+100);
    configurator.columnWidth(1, defaultColumnWidth+200);
    configurator.mergeCell(0, 0, 0, 1);
    configurator.defaultRowHeight((short)(defaultRowHeight+100));
    configurator.rightToLeft(!isRightToLeft);
    configurator.rowSumsRight(!isRowSumsRight);
    configurator.rowSumsBelow(!isRowSumsBellow);
    configurator.selected(!isSelected);

    //then
    boolean expectedIsAutoBreaks = !autoBreaks;
    Assertions.assertEquals(expectedIsAutoBreaks, sheet.getAutobreaks());

    int expectedActiveCellRow = 0;
    int expectedActiveCellCol = 0;
    Assertions.assertEquals(expectedActiveCellRow, sheet.getActiveCell().getRow());
    Assertions.assertEquals(expectedActiveCellCol, sheet.getActiveCell().getColumn());

    int expectedColumnWidth0 = defaultColumnWidth+100;
    int expectedColumnWidth1 = defaultColumnWidth+200;
    Assertions.assertEquals(expectedColumnWidth0, sheet.getColumnWidth(0));
    Assertions.assertEquals(expectedColumnWidth1, sheet.getColumnWidth(1));

    Assertions.assertNotNull(
        sheet.getMergedRegions()
            .stream()
            .filter(cellAddresses -> cellAddresses.isInRange(0, 1))
            .findFirst()
    );

    short expectedDefaultRowHeight = (short)(defaultRowHeight+100);
    Assertions.assertEquals(expectedDefaultRowHeight, sheet.getDefaultRowHeight());

    boolean expectedIsRightToLeft = !isRightToLeft;
    Assertions.assertEquals(expectedIsRightToLeft, sheet.isRightToLeft());

    boolean expectedIsRowSumsRight = !isRowSumsRight;
    Assertions.assertEquals(expectedIsRowSumsRight, sheet.getRowSumsRight());

    boolean expectedIsRowSumsBellow = !isRowSumsBellow;
    Assertions.assertEquals(expectedIsRowSumsBellow, sheet.getRowSumsBelow());

    boolean expectedIsSelected = !isSelected;
    Assertions.assertEquals(expectedIsSelected, sheet.isSelected());
  }

  @Test
  @Order(2)
  @DisplayName("프린트 세부 설정 테스트")
  public void printOptionsTest(){
    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet();
    SheetConfigurer configurator = new SheetConfigurer(sheet);

    boolean isHorizontallyCenter = sheet.getHorizontallyCenter();
    boolean isVerticallyCenter = sheet.getVerticallyCenter();
    int[] columnBreaks = sheet.getColumnBreaks();
    double marginTop = sheet.getMargin(PageMargin.TOP);
    double marginLeft = sheet.getMargin(PageMargin.LEFT);
    double marginBottom = sheet.getMargin(PageMargin.BOTTOM);
    double marginRight = sheet.getMargin(PageMargin.RIGHT);
    double marginHeader = sheet.getMargin(PageMargin.HEADER);
    double marginFooter = sheet.getMargin(PageMargin.FOOTER);
    boolean isPrintGridlines = sheet.isPrintGridlines();
    boolean isPrintRowAndColumnHeadings = sheet.isPrintRowAndColumnHeadings();
    boolean isDisplayGuts = sheet.getDisplayGuts();
    boolean isDisplayGridlines = sheet.isDisplayGridlines();
    boolean isDisplayFormulas = sheet.isDisplayFormulas();
    boolean isForceFormulaRecalculation = sheet.getForceFormulaRecalculation();
    boolean isDisplayRowColHeadings = sheet.isDisplayRowColHeadings();
    boolean isDisplayZeros = sheet.isDisplayZeros();

    //when
    configurator.horizontallyCenter(!isHorizontallyCenter);
    configurator.verticallyCenter(!isVerticallyCenter);
    configurator.columnBreak(100);
    configurator.columnBreak(200);
    configurator.marginTop(marginTop+10);
    configurator.marginLeft(marginLeft+20);
    configurator.marginBottom(marginBottom+30);
    configurator.marginRight(marginRight+40);
    configurator.marginHeader(marginHeader+50);
    configurator.marginFooter(marginFooter+60);
    configurator.printGridlines(!isPrintGridlines);
    configurator.printRowAndColumnHeadings(!isPrintRowAndColumnHeadings);
    configurator.repeatingColumns(2, 3, 2, 3);
    configurator.repeatingRows(2, 3, 2, 3);
    configurator.rowBreaks(15);
    configurator.rowBreaks(30);
    configurator.displayGuts(!isDisplayGuts);
    configurator.displayGridlines(!isDisplayGridlines);
    configurator.displayFormulas(!isDisplayFormulas);
    configurator.forceFormulaRecalculation(!isForceFormulaRecalculation);
    configurator.displayRowColHeadings(!isDisplayRowColHeadings);
    configurator.displayZeros(!isDisplayZeros);

    //then
    boolean expectedIsHorizontallyCenter = !isHorizontallyCenter;
    Assertions.assertEquals(expectedIsHorizontallyCenter, sheet.getHorizontallyCenter());

    boolean expectedIsVerticallyCenter = !isVerticallyCenter;
    Assertions.assertEquals(expectedIsVerticallyCenter, sheet.getVerticallyCenter());

    int expectedColumnBreak1 = 100;
    int expectedColumnBreak2 = 200;
    Assertions.assertEquals(expectedColumnBreak1, sheet.getColumnBreaks()[0]);
    Assertions.assertEquals(expectedColumnBreak2, sheet.getColumnBreaks()[1]);

    double expectedMarginTop = marginTop+10;
    double expectedMarginLeft = marginLeft+20;
    double expectedMarginBottom = marginBottom+30;
    double expectedMarginRight = marginRight+40;
    double expectedMarginHeader = marginHeader+50;
    double expectedMarginFooter = marginFooter+60;
    Assertions.assertEquals(expectedMarginTop, sheet.getMargin(PageMargin.TOP));
    Assertions.assertEquals(expectedMarginLeft, sheet.getMargin(PageMargin.LEFT));
    Assertions.assertEquals(expectedMarginBottom, sheet.getMargin(PageMargin.BOTTOM));
    Assertions.assertEquals(expectedMarginRight, sheet.getMargin(PageMargin.RIGHT));
    Assertions.assertEquals(expectedMarginHeader, sheet.getMargin(PageMargin.HEADER));
    Assertions.assertEquals(expectedMarginFooter, sheet.getMargin(PageMargin.FOOTER));

    boolean expectedIsPrintGridlines = !isPrintGridlines;
    boolean expectedIsPrintRowAndColumnHeadings = !isPrintRowAndColumnHeadings;
    Assertions.assertEquals(expectedIsPrintGridlines, sheet.isPrintGridlines());
    Assertions.assertEquals(expectedIsPrintRowAndColumnHeadings, sheet.isPrintRowAndColumnHeadings());

    int expectedRepeatingRowsStartRow = 2;
    int expectedRepeatingRowsEndRow = 3;
    CellRangeAddress repeatingRows = sheet.getRepeatingRows();
    Assertions.assertEquals(expectedRepeatingRowsStartRow, repeatingRows.getFirstRow());
    Assertions.assertEquals(expectedRepeatingRowsEndRow, repeatingRows.getLastRow());

    int expectedRepeatingColumnsStartCol = 2;
    int expectedRepeatingColumnsEndCol = 3;
    CellRangeAddress repeatingColumns = sheet.getRepeatingColumns();
    Assertions.assertEquals(expectedRepeatingColumnsStartCol, repeatingColumns.getFirstColumn());
    Assertions.assertEquals(expectedRepeatingColumnsEndCol, repeatingColumns.getLastColumn());

    int expectedRowBreaks1 = 15;
    int expectedRowBreaks2 = 30;
    Assertions.assertEquals(expectedRowBreaks1, sheet.getRowBreaks()[0]);
    Assertions.assertEquals(expectedRowBreaks2, sheet.getRowBreaks()[1]);

    boolean expectedDisplayGuts = !isDisplayGuts;
    Assertions.assertEquals(expectedDisplayGuts, sheet.getDisplayGuts());

    boolean expectedIsDisplayGridlines = !isDisplayGridlines;
    Assertions.assertEquals(expectedIsDisplayGridlines, sheet.isDisplayGridlines());

    boolean expectedIsDisplayFormulas = !isDisplayFormulas;
    Assertions.assertEquals(expectedIsDisplayFormulas, sheet.isDisplayFormulas());

    boolean expectedIsForceFormulaRecalculation = !isForceFormulaRecalculation;
    Assertions.assertEquals(expectedIsForceFormulaRecalculation, sheet.getForceFormulaRecalculation());

    boolean expectedIsDisplayZeros = !isDisplayZeros;
    Assertions.assertEquals(expectedIsDisplayZeros, sheet.isDisplayZeros());

    boolean expectedIsDisplayRowColHeadings = !isDisplayRowColHeadings;
    Assertions.assertEquals(expectedIsDisplayRowColHeadings, sheet.isDisplayRowColHeadings());

  }
}
