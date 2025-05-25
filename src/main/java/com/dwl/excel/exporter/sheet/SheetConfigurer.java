package com.dwl.excel.exporter.sheet;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public final class SheetConfigurer {

  private final Sheet sheet;

  public SheetConfigurer(Sheet sheet){
    this.sheet = sheet;
  }

  /**
   * 셀 병합 영역 설정
   * @param startRow
   * @param endRow
   * @param startColumn
   * @param endColumn
   * @return
   */
  public SheetConfigurer mergeCell(int startRow, int endRow, int startColumn, int endColumn){
    this.sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startColumn, endColumn));
    return this;
  }

  /**
   * 자동 페이지 나누기 설정
   * @param isAutoBreak
   * @return
   */
  public SheetConfigurer autoBreaks(boolean isAutoBreak){
    this.sheet.setAutobreaks(isAutoBreak);
    return this;
  }

  /**
   * 지정한 열에 드롭다운 필터를 자동으로 추가
   * @param startRow
   * @param endRow
   * @param startColumn
   * @param endColumn
   * @return
   */
  public SheetConfigurer autoFilter(int startRow, int endRow, int startColumn, int endColumn){
    this.sheet.setAutoFilter(new CellRangeAddress(startRow, endRow, startColumn, endColumn));
    return this;
  }

  /**
   * 워크시트를 열었을 때 커서가 위치할 기본 셀을 지정
   * @param row
   * @param column
   * @return
   */
  public SheetConfigurer activeCell(int row, int column){
    this.sheet.setActiveCell(new CellAddress(row, column));
    return this;
  }

  /**
   * 특정 열에 자동크기조정 설정
   * @param columnIndex
   * @return
   */
  public SheetConfigurer autoSizeColumn(int columnIndex){
    this.sheet.autoSizeColumn(columnIndex);
    return this;
  }

  /**
   * 특정 병합 열에 자동크기조정 설정
   * @param columnIndex
   * @return
   */
  public SheetConfigurer autoSizeColumn(int columnIndex, boolean useMergedCell){
    this.sheet.autoSizeColumn(columnIndex, useMergedCell);
    return this;
  }

  /**
   * 열 그룹 설정
   * @param fromColumn
   * @param toColumn
   * @return
   */
  public SheetConfigurer groupColumn(int fromColumn, int toColumn){
    this.sheet.groupColumn(fromColumn, toColumn);
    return this;
  }

  /**
   * 열 그룹을 접거나 펼칠지 지정 (그룹화된 열)
   * @param column
   * @param collapsed
   * @return
   */
  public SheetConfigurer columnGroupCollapsed(int column, boolean collapsed){
    this.sheet.setColumnGroupCollapsed(column, collapsed);
    return this;
  }

  /**
   * 특정 열을 숨기거나 보이게 설정
   * @param column
   * @param hidden
   * @return
   */
  public SheetConfigurer columnHidden(int column, boolean hidden){
    this.sheet.setColumnHidden(column, hidden);
    return this;
  }

  /**
   * 전체 시트의 기본 열 너비를 설정
   * @param width
   * @return
   */
  public SheetConfigurer defaultColumnWidth(int width){
    this.sheet.setDefaultColumnWidth(width);
    return this;
  }

  /**
   * 열의 너비를 설정 (단위는 1/256 character width)
   * @param column
   * @param width
   * @return
   */
  public SheetConfigurer columnWidth(int column, int width){
    this.sheet.setColumnWidth(column, width);
    return this;
  }

  /**
   * 전체 시트의 기본 행 높이 (단위: 1/20 포인트)를 설정합니다.
   * @param height
   * @return
   */
  public SheetConfigurer defaultRowHeight(short height){
    this.sheet.setDefaultRowHeight(height);
    return this;
  }

  /**
   * 전체 시트의 기본 행 높이 (포인트 단위) 를 설정
   * @param height
   * @return
   */
  public SheetConfigurer defaultRowHeightInPoints(float height){
    this.sheet.setDefaultRowHeightInPoints(height);
    return this;
  }

  /**
   * 시트를 오른쪽에서 왼쪽으로 방향 전환
   * @param conversion
   * @return
   */
  public SheetConfigurer rightToLeft(boolean conversion){
    this.sheet.setRightToLeft(conversion);
    return this;
  }

  /**
   * 열 그룹 설정
   * @param fromRow
   * @param toRow
   * @return
   */
  public SheetConfigurer groupRow(int fromRow, int toRow){
    this.sheet.groupRow(fromRow, toRow);
    return this;
  }


  /**
   * 그룹화된 행을 접거나 펼치기 설정
   * @param row
   * @param collapsed
   * @return
   */
  public SheetConfigurer rowGroupCollapsed(int row, boolean collapsed){
    this.sheet.setRowGroupCollapsed(row, collapsed);
    return this;
  }

  /**
   * 개요 모드에서 합계 행을 그룹 아래에 표시할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer rowSumsBelow(boolean display){
    this.sheet.setRowSumsBelow(display);
    return this;
  }

  /**
   * 개요 모드에서 합계 열을 오른쪽에 표시할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer rowSumsRight(boolean display){
    this.sheet.setRowSumsRight(display);
    return this;
  }

  /**
   * 이 워크시트를 기본 선택된 시트로 표시
   * @param selected
   * @return
   */
  public SheetConfigurer selected(boolean selected){
    this.sheet.setSelected(selected);
    return this;
  }

  /**
   * 시트의 보기 배율 설정 (예: 100 = 100%)
   * @param rate
   * @return
   */
  public SheetConfigurer zoom(int rate){
    this.sheet.setZoom(rate);
    return this;
  }

  /**
   * 인쇄 시 내용을 가로 방향으로 중앙 정렬할지 설정
   * @param sorted
   * @return
   */
  public SheetConfigurer horizontallyCenter(boolean sorted){
    this.sheet.setHorizontallyCenter(sorted);
    return this;
  }

  /**
   * 인쇄 시 세로 방향으로 중앙 정렬할지 설정
   * @param sorted
   * @return
   */
  public SheetConfigurer verticallyCenter(boolean sorted){
    this.sheet.setVerticallyCenter(sorted);
    return this;
  }

  /**
   * 특정 열에서 인쇄 시 강제 페이지 나눔을 설정합니다.
   * @param column
   * @return
   */
  public SheetConfigurer columnBreak(int column){
    this.sheet.setColumnBreak(column);
    return this;
  }

  /**
   * 특정 열에서 인쇄 시 강제 페이지 나눔을 설정합니다.
   * @param fit
   * @return
   */
  public SheetConfigurer fitToPage(boolean fit){
    this.sheet.setFitToPage(fit);
    return this;
  }

  /**
   * 인쇄 시 페이지 TOP 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginTop(double margin){
    this.sheet.setMargin(PageMargin.TOP, margin);
    return this;
  }

  /**
   * 인쇄 시 페이지 LEFT 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginLeft(double margin){
    this.sheet.setMargin(PageMargin.LEFT, margin);
    return this;
  }

  /**
   * 인쇄 시 페이지 BOTTOM 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginBottom(double margin){
    this.sheet.setMargin(PageMargin.BOTTOM, margin);
    return this;
  }

  /**
   * 인쇄 시 페이지 RIGHT 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginRight(double margin){
    this.sheet.setMargin(PageMargin.RIGHT, margin);
    return this;
  }

  /**
   * 인쇄 시 페이지 HEADER 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginHeader(double margin){
    this.sheet.setMargin(PageMargin.HEADER, margin);
    return this;
  }

  /**
   * 인쇄 시 페이지 FOOTER 여백 설정
   * @param margin
   * @return
   */
  public SheetConfigurer marginFooter(double margin){
    this.sheet.setMargin(PageMargin.FOOTER, margin);
    return this;
  }

  /**
   * 인쇄 시 셀의 격자선을 포함할지 설정
   * @param print
   * @return
   */
  public SheetConfigurer printGridlines(boolean print){
    this.sheet.setPrintGridlines(print);
    return this;
  }

  /**
   * 인쇄 시 행 번호/열 문자를 포함할지 설정
   * @param print
   * @return
   */
  public SheetConfigurer printRowAndColumnHeadings(boolean print){
    this.sheet.setPrintRowAndColumnHeadings(print);
    return this;
  }

  /**
   * 인쇄할 때 각 페이지에 반복 출력할 열을 지정
   * @param startRow
   * @param endRow
   * @param startCol
   * @param endCol
   * @return
   */
  public SheetConfigurer repeatingColumns(int startRow, int endRow, int startCol, int endCol){
    this.sheet.setRepeatingColumns(new CellRangeAddress(startRow, endRow, startCol, endCol));
    return this;
  }

  /**
   * 인쇄할 때 각 페이지에 반복 출력할 행을 지정
   * @param startRow
   * @param endRow
   * @param startCol
   * @param endCol
   * @return
   */
  public SheetConfigurer repeatingRows(int startRow, int endRow, int startCol, int endCol){
    this.sheet.setRepeatingRows(new CellRangeAddress(startRow, endRow, startCol, endCol));
    return this;
  }

  /**
   * 특정 행에서 강제 인쇄 페이지 나눔 설정
   * @param row
   * @return
   */
  public SheetConfigurer rowBreaks(int row){
    this.sheet.setRowBreak(row);
    return this;
  }

  /**
   * 셀 범위에 배열 수식을 설정 (예: {=A1:A10*B1:B10})
   * @param s
   * @param startRow
   * @param endRow
   * @param startCol
   * @param endCol
   * @return
   */
  public SheetConfigurer arrayFormula(String s, int startRow, int endRow, int startCol, int endCol){
    this.sheet.setArrayFormula(s, new CellRangeAddress(startRow, endRow, startCol, endCol));
    return this;
  }

  /**
   * 수식 셀을 결과가 아니라 수식 그대로 보이게 할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer displayFormulas(boolean display){
    this.sheet.setDisplayFormulas(display);
    return this;
  }

  /**
   * 파일을 열 때 수식을 강제로 다시 계산하도록 설정
   */
  public SheetConfigurer forceFormulaRecalculation(boolean recalculation){
    this.sheet.setForceFormulaRecalculation(recalculation);
    return this;
  }

  /**
   * 엑셀 편집기에서 셀 테두리(격자선)를 표시할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer displayGridlines(boolean display){
    this.sheet.setDisplayGridlines(display);
    return this;
  }

  /**
   * 그룹화된 행/열의 개요 표시 기호(+)를 표시할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer displayGuts(boolean display){
    this.sheet.setDisplayGuts(display);
    return this;
  }

  /**
   * 좌측/상단의 행 번호 및 열 문자(A, B, C...)를 표시할지 설정
   * @param display
   * @return
   */
  public SheetConfigurer displayRowColHeadings(boolean display){
    this.sheet.setDisplayRowColHeadings(display);
    return this;
  }

  /**
   * 값이 0인 셀에 0을 표시할지 말지 설정
   * @param display
   * @return
   */
  public SheetConfigurer displayZeros(boolean display){
    this.sheet.setDisplayZeros(display);
    return this;
  }
}
